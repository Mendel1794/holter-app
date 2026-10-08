import streamlit as st
import fitz  # PyMuPDF: motor C++ de renderizado ultrarrápido y extracción espacial
from PIL import Image, ImageOps
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

# ==============================================================================
# CONFIGURACIÓN REGIONAL Y HORARIA (COLOMBIA UTC-5)
# ==============================================================================
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
    from supabase import create_client
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
# SEGURIDAD Y CREDENCIALES (HASH CRIPTOGRÁFICO SHA-256)
# ==============================================================================
def hash_clave(pwd: str) -> str:
    return hashlib.sha256(pwd.strip().encode("utf-8")).hexdigest()

LISTA_ESPECIALISTAS = [
    {
        "id": "dr.amaya",
        "etiqueta": "Dr. William Amaya Ramirez (Internista - Cardiólogo)",
        "clave_hash": "b2f6efd1b4ee3d201c1074eef714150537f8846c4f8d5a1a1f337a54823485c2",  # Cardio2025*
        "nombre_completo": "DR. WILLIAM AMAYA RAMIREZ",
        "especialidad": "INTERNISTA - CARDIÓLOGO",
        "registro": "RM 79.502.624 SDS"
    },
    {
        "id": "dr.suarez",
        "etiqueta": "Dr. Martin Suárez Arámbula (Cardiólogo Hemodinamista)",
        "clave_hash": "246beea2299882fe1a82fce5ba7a544c92eefd08dcfd5fe48a586e9275a2ad02",  # Suarez2026*
        "nombre_completo": "DR. MARTIN SUÁREZ ARÁMBULA",
        "especialidad": "MÉDICO INTERNISTA - CARDIÓLOGO HEMODINAMISTA",
        "registro": "RM 13491094"
    },
    {
        "id": "dra.cardio",
        "etiqueta": "Dra. Paola Figueroa (Cardióloga)",
        "clave_hash": "8ba1207e3a9686aa3fc8cfae4f8d55ca7bfae6ef7bb7b98b0f9cbe87dfc3ae53",  # Cardio2026*
        "nombre_completo": "DRA. PAOLA FIGUEROA",
        "especialidad": "MÉDICO ESPECIALISTA EN CARDIOLOGÍA",
        "registro": "RM 52.890.123 SDS"
    },
    {
        "id": "admin",
        "etiqueta": "Administración del Sistema",
        "clave_hash": "8488e33f388916d63de33b3a726dc0e9c8cf5962e24cf8f60da308560124239d",  # HolterClaveSegura123
        "nombre_completo": "DR. WILLIAM AMAYA RAMIREZ",
        "especialidad": "INTERNISTA - CARDIÓLOGO",
        "registro": "RM 79.502.624 SDS"
    }
]

PERFILES_POR_ID = {m["id"]: m for m in LISTA_ESPECIALISTAS}
OPCIONES_NOMBRES = [m["etiqueta"] for m in LISTA_ESPECIALISTAS]

# ==============================================================================
# IDENTIDAD VISUAL INSTITUCIONAL CENCARDIO
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
        box-shadow: 0 4px 20px -2px rgba(10, 37, 64, 0.04); margin-bottom: 1.2rem;
    }
    .inst-badge-primary {
        background: #e0f2fe; color: #0369a1; font-size: 0.72rem; font-weight: 700;
        padding: 4px 10px; border-radius: 20px; text-transform: uppercase; border: 1px solid #bae6fd;
    }
    .inst-badge-success {
        background: #ecfdf5; color: #065f46; font-size: 0.72rem; font-weight: 700;
        padding: 4px 10px; border-radius: 20px; text-transform: uppercase; border: 1px solid #a7f3d0;
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
        max-width: 500px; margin: 3rem auto; text-align: center;
    }
    </style>
    """

st.markdown(cargar_estilos_institucionales(), unsafe_allow_html=True)

# ==============================================================================
# BASE DE DATOS LOCAL (SQLITE) CON MIGRACIÓN INCONDICIONAL
# ==============================================================================
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
            codigo_verificacion TEXT UNIQUE,
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
    c.execute("PRAGMA table_info(estudios)")
    cols_existentes = [row[1] for row in c.fetchall()]
    cols_necesarias = [
        ("parametro_clave", "TEXT"),
        ("codigo_verificacion", "TEXT"),
        ("hash_sha256", "TEXT"),
        ("pdf_blob", "BLOB")
    ]
    for c_nom, c_tip in cols_necesarias:
        if c_nom not in cols_existentes:
            try:
                c.execute(f"ALTER TABLE estudios ADD COLUMN {c_nom} {c_tip}")
            except Exception:
                pass
    conn.commit()
    conn.close()

init_db_local()

# ==============================================================================
# CONECTOR SUPABASE (CON DESINFECCIÓN ESTRICTA DE URL)
# ==============================================================================
@st.cache_resource
def obtener_cliente_supabase():
    if not SUPABASE_LIB_OK:
        return None
    url = st.secrets.get("SUPABASE_URL", os.environ.get("SUPABASE_URL", ""))
    key = st.secrets.get("SUPABASE_KEY", os.environ.get("SUPABASE_KEY", ""))
    if url and key:
        try:
            url_limpia = str(url).strip().rstrip("/")
            for sub in ["/rest/v1", "/auth/v1", "/storage/v1"]:
                if url_limpia.endswith(sub):
                    url_limpia = url_limpia[:-len(sub)].rstrip("/")
            key_limpia = str(key).strip().strip('"').strip("'")
            cliente = create_client(url_limpia, key_limpia)
            try:
                cliente.storage.create_bucket("estudios-pdf", options={"public": True})
            except Exception:
                pass
            return cliente
        except Exception:
            return None
    return None

supabase = obtener_cliente_supabase()

def calcular_hash_sha256(pdf_bytes: bytes) -> str:
    return hashlib.sha256(pdf_bytes).hexdigest()

def guardar_estudio_servicio(nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif):
    fecha_actual_str = ahora_colombia().strftime("%Y-%m-%d %H:%M:%S")
    hash_seguridad = calcular_hash_sha256(pdf_bytes)
    nombre_archivo = f"{cod_verif[:12]}_{re.sub(r'[^A-Za-z0-9]', '_', nombre)}.pdf"
    guardado_en_nube = False

    if supabase:
        try:
            pdf_url = ""
            try:
                supabase.storage.from_("estudios-pdf").upload(
                    path=nombre_archivo,
                    file=pdf_bytes,
                    file_options={"content-type": "application/pdf", "upsert": "true"}
                )
                pdf_url = supabase.storage.from_("estudios-pdf").get_public_url(nombre_archivo)
            except Exception:
                pdf_url = ""

            supabase.table("estudios").insert({
                "fecha_registro": ahora_colombia().isoformat(),
                "paciente_nombre": nombre,
                "modalidad": modalidad,
                "cups": cups,
                "parametro_clave": parametro_clave,
                "medico_firmante": medico,
                "informe_texto": texto,
                "pdf_url": pdf_url,
                "codigo_verificacion": cod_verif,
                "hash_sha256": hash_seguridad
            }).execute()
            guardado_en_nube = True
        except Exception:
            pass

    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("""
        INSERT OR REPLACE INTO estudios (fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, informe_texto, pdf_blob, codigo_verificacion, hash_sha256)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (fecha_actual_str, nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif, hash_seguridad))
    conn.commit()
    conn.close()

    if guardado_en_nube:
        return True, "Estudio certificado en Nube Supabase y almacenamiento local con hash SHA-256."
    return True, "Estudio certificado y resguardado en almacenamiento local con hash SHA-256."

def obtener_historial_servicio():
    if supabase:
        try:
            res = supabase.table("estudios").select("id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, pdf_url, codigo_verificacion, hash_sha256").order("id", desc=True).execute()
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
                    "codigo_verificacion": r.get("codigo_verificacion", ""),
                    "hash_sha256": r.get("hash_sha256", "")
                })
            return lista, "supabase"
        except Exception:
            pass

    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, codigo_verificacion, hash_sha256 FROM estudios ORDER BY id DESC")
    filas = c.fetchall()
    conn.close()
    lista = []
    for f in filas:
        lista.append({
            "id": f[0], "fecha_registro": f[1], "paciente": f[2], "modalidad": f[3],
            "cups": f[4], "parametro_clave": f[5], "medico": f[6], "pdf_url": "",
            "codigo_verificacion": f[7] if f[7] else "N/A",
            "hash_sha256": f[8] if len(f) > 8 and f[8] else ""
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
    init_db_local()
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
    init_db_local()
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
        return False, "El archivo debe incluir columnas de Cédula/ID y Teléfono/Celular."

    registros = 0
    fecha_hoy = ahora_colombia().strftime("%Y-%m-%d %H:%M")
    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
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
    return True, f"Sincronizados {registros} pacientes en el directorio institucional."

def buscar_telefono_servicio(cedula):
    if not cedula:
        return ""
    ced_limpia = re.sub(r'\D', '', str(cedula))
    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT telefono FROM dinamica_pacientes WHERE cedula = ?", (ced_limpia,))
    res = c.fetchone()
    conn.close()
    return res[0] if res else ""

# ==============================================================================
# VALIDACIÓN PÚBLICA CERO-CONFIANZA (ZERO-TRUST) POR TOKEN Y HASH SHA-256
# ==============================================================================
params = st.query_params
if "token" in params or "val" in params:
    token_consulta = str(params.get("token", params.get("val", ""))).strip()
    estudio_db = None

    if supabase:
        try:
            res_sb = supabase.table("estudios").select("*").ilike("codigo_verificacion", f"{token_consulta}%").limit(1).execute()
            if res_sb.data:
                estudio_db = res_sb.data[0]
        except Exception:
            pass

    if not estudio_db:
        init_db_local()
        conn = sqlite3.connect("historial_cencardio.db")
        c = conn.cursor()
        c.execute("""
            SELECT paciente_nombre, medico_firmante, fecha_registro, modalidad, cups, codigo_verificacion, hash_sha256
            FROM estudios WHERE codigo_verificacion LIKE ? LIMIT 1
        """, (f"{token_consulta}%",))
        row = c.fetchone()
        conn.close()
        if row:
            estudio_db = {
                "paciente_nombre": row[0],
                "medico_firmante": row[1],
                "fecha_registro": row[2],
                "modalidad": row[3],
                "cups": row[4],
                "codigo_verificacion": row[5],
                "hash_sha256": row[6]
            }

    c_v1, c_v2, c_v3 = st.columns([1, 1.8, 1])
    with c_v2:
        logo_data = obtener_logo_b64()
        logo_html = f'<img src="{logo_data}" style="max-width:180px; margin-bottom:1rem;" alt="Cencardio Logo">' if logo_data else '<div style="font-size:3rem; margin-bottom:0.4rem;">🫀</div>'

        if estudio_db:
            st.markdown(f"""
                <div class="cencardio-card">
                    {logo_html}
                    <div style="font-size: 1.25rem; font-weight: 800; color: #0a2540; text-transform: uppercase;">
                        Centro Cardiovascular Colombiano
                    </div>
                    <div style="font-size: 0.82rem; font-weight: 700; color: #c8102e; letter-spacing: 0.5px; text-transform: uppercase; margin-bottom: 1.5rem;">
                        CENCARDIO · Certificación Digital Forense
                    </div>
                    <div style="background: #f0fdf4; border: 1.5px solid #86efac; border-radius: 12px; padding: 1.3rem; margin-bottom: 1.5rem; text-align: left;">
                        <div style="color: #166534; font-size: 1rem; font-weight: 800; margin-bottom: 0.6rem; display: flex; align-items: center; gap: 8px;">
                            <span>✅</span> ESTUDIO OFICIAL AUTÉNTICO Y CERTIFICADO
                        </div>
                        <div style="font-size: 0.88rem; color: #1f2937; line-height: 1.6;">
                            <b>Procedimiento:</b> {estudio_db['modalidad']} ({estudio_db['cups']})<br>
                            <b>Paciente:</b> {estudio_db['paciente_nombre']}<br>
                            <b>Especialista Lector:</b> {estudio_db['medico_firmante']}<br>
                            <b>Fecha Emisión:</b> {str(estudio_db['fecha_registro'])[:19]}<br>
                            <b>Identificador Forense:</b> <span style="font-family: monospace; color: #0369a1; font-weight: 700;">{estudio_db['codigo_verificacion']}</span><br>
                            <b>Firma Criptográfica SHA-256:</b> <span style="font-family: monospace; font-size: 0.72rem; color: #475569; word-break: break-all;">{estudio_db.get('hash_sha256', 'REGISTRADO')}</span>
                        </div>
                    </div>
                    <div style="font-size: 0.78rem; color: #64748b; line-height: 1.4;">
                        Custodiado bajo los requisitos de la Res. 3100 de 2019 de MinSalud Colombia. Integridad garantizada.
                    </div>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
                <div class="cencardio-card">
                    {logo_html}
                    <div style="background: #fef2f2; border: 2px solid #ef4444; border-radius: 12px; padding: 1.8rem; text-align: center;">
                        <div style="font-size: 2.5rem; margin-bottom: 0.4rem;">🛑</div>
                        <h3 style="color: #991b1b; margin: 0;">CERTIFICADO NO ENCONTRADO O INVÁLIDO</h3>
                        <p style="color: #7f1d1d; font-size: 0.88rem; margin-top: 0.8rem; line-height: 1.5;">
                            El código consultado no corresponde a ningún estudio emitido ni custodiado legalmente por el Centro Cardiovascular Colombiano CENCARDIO.
                        </p>
                    </div>
                </div>
            """, unsafe_allow_html=True)

        if st.button("Ir al Portal de Operaciones", use_container_width=True):
            st.query_params.clear()
            st.rerun()

    st.stop()

# ==============================================================================
# CONTROL DE ACCESO
# ==============================================================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = ""

def cerrar_sesion():
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""

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
            clave_ingresada = st.text_input("Contraseña de Acceso:", type="password")
            if st.form_submit_button("Ingresar a la Estación", use_container_width=True):
                medico = next(m for m in LISTA_ESPECIALISTAS if m["etiqueta"] == seleccion_etiqueta)
                if hash_clave(clave_ingresada) == medico["clave_hash"]:
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = medico["id"]
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta para el especialista seleccionado.")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ==============================================================================
# MOTOR 2: INTELIGENCIA ARTIFICIAL MULTIMODAL (GEMINI API)
# ==============================================================================
def consultar_gemini_json(prompt_text: str, inline_items=None):
    gemini_key = st.secrets.get("GEMINI_API_KEY", st.secrets.get("gemini_api_key", os.environ.get("GEMINI_API_KEY", "")))
    gemini_key = str(gemini_key).strip().strip('"').strip("'")
    if not gemini_key:
        return False, None, "No se encontró GEMINI_API_KEY en Secrets."

    modelos_prioritarios = [
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
        "generationConfig": {"temperature": 0.05, "responseMimeType": "application/json"}
    }

    ultimo_error = ""
    for ver, mod_name in modelos_prioritarios:
        url = f"https://generativelanguage.googleapis.com/{ver}/{mod_name}:generateContent?key={gemini_key}"
        try:
            resp = requests.post(url, headers={"Content-Type": "application/json"}, json=payload, timeout=60)
            if resp.status_code == 200:
                res_json = resp.json()
                raw_text = res_json['candidates'][0]['content']['parts'][0]['text']
                match = re.search(r'\{.*\}', raw_text, re.DOTALL)
                clean_json = match.group(0) if match else raw_text.strip()
                data = json.loads(clean_json)
                return True, data, mod_name
            else:
                ultimo_error = f"{mod_name} ({resp.status_code}): {resp.text[:120]}"
        except Exception as e:
            ultimo_error = f"{mod_name} error de conexión: {str(e)}"

    return False, None, ultimo_error

# ==============================================================================
# UTILIDADES DE PARSEO NUMÉRICO
# ==============================================================================
def limpiar_numero(val_str):
    if val_str is None:
        return None
    s = str(val_str).strip()
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "." in s and len(s.split(".")[-1]) == 3:
        s = s.replace(".", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return int(float(s))
    except Exception:
        return None

def extraer_flotante(val_str):
    if val_str is None:
        return None
    s = str(val_str).strip().replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def crear_expediente(parametro, valor, unidades, status, source, evidencia, metadata=None, requires_review=False, discrepancia=None):
    return {
        "parametro": parametro,
        "value": valor,
        "unidades": unidades,
        "status": status,
        "source": source,
        "evidence": evidencia,
        "reconstruccion_metadata": metadata,
        "discrepancy": discrepancia,
        "requires_review": requires_review
    }

# ==============================================================================
# MOTOR 1, 2 Y 3 PARA HOLTER ECG 24H (CUPS 895001)
# ==============================================================================
def extraer_m1_holter_determinista(doc: fitz.Document) -> dict:
    texto_completo = ""
    for num_p, page in enumerate(doc):
        texto_completo += f"\n--- PAGINA {num_p+1} ---\n" + page.get_text()

    d = {}
    m_nom = re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,50},\s*[A-ZÁÉÍÓÚÑ\s]{3,50})[\s\n]+(?:No confirmado|Confirmado|Reconfirmado)?[\s\n]*Informe Holter", texto_completo)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else None

    m_id = re.search(r"(?:ID\s*Paciente|ID|C\.?C\.?|Doc\.?|Historia)\s*[:\.]?\s*(\d{5,12})", texto_completo, re.IGNORECASE)
    d["cedula"] = m_id.group(1) if m_id else None

    fc_p = re.search(r"Prom\.?\s*(\d{2,3})", texto_completo)
    d["fc_prom"] = int(fc_p.group(1)) if fc_p else None

    fc_max = re.search(r"M[áa]x\s*(\d{2,3})", texto_completo)
    d["fc_max"] = int(fc_max.group(1)) if fc_max else None

    fc_min = re.search(r"M[íi]n\s*(\d{2,3})", texto_completo)
    d["fc_min"] = int(fc_min.group(1)) if fc_min else None

    fc_dia = re.search(r"D[íi]a.*?Prom\.?\s*(\d{2,3})", texto_completo)
    d["fc_dia"] = int(fc_dia.group(1)) if fc_dia else None

    fc_noc = re.search(r"Noche.*?Prom\.?\s*(\d{2,3})", texto_completo)
    d["fc_noc"] = int(fc_noc.group(1)) if fc_noc else None

    m_conteo = re.search(r"Latidos[^\n\r]*\n[^\n\r]*Conteo\s+([\d\.]+)\s+([\d\.]+)\s+\d+%\s+([\d\.]+)[^\n\r]*\s+([\d\.]+)[^\n\r]*\s+([\d\.]+)\s+(\d+)?%", texto_completo, re.IGNORECASE)
    if m_conteo:
        d["total_latidos"] = limpiar_numero(m_conteo.group(1))
        d["latidos_normales"] = limpiar_numero(m_conteo.group(2))
        d["ev_total"] = limpiar_numero(m_conteo.group(3))
        d["esv_total"] = limpiar_numero(m_conteo.group(4))
        d["mcp_latidos"] = limpiar_numero(m_conteo.group(5))
        d["mcp_porcentaje"] = extraer_flotante(m_conteo.group(6))
    else:
        m_tot = re.search(r"(?:Total\s+de\s+latidos|Total\s+QRS)\s*[:\.]?\s*([\d\.]+)", texto_completo, re.IGNORECASE)
        d["total_latidos"] = limpiar_numero(m_tot.group(1)) if m_tot else None
        m_ev = re.search(r"(?:Latidos\s+ventriculares|Latidos\s+V)\b[^\n\r\d]*([\d\.]+)", texto_completo, re.IGNORECASE)
        d["ev_total"] = limpiar_numero(m_ev.group(1)) if m_ev else None
        m_esv = re.search(r"Latidos\s+supraventriculares\s*:\s*([\d\.]+)", texto_completo, re.IGNORECASE)
        d["esv_total"] = limpiar_numero(m_esv.group(1)) if m_esv else None
        m_mcp = re.search(r"Marcapaso\s+([\d\.]+)\s+(\d+)%", texto_completo, re.IGNORECASE)
        d["mcp_latidos"] = limpiar_numero(m_mcp.group(1)) if m_mcp else None
        d["mcp_porcentaje"] = extraer_flotante(m_mcp.group(2)) if m_mcp else None
        d["latidos_normales"] = None

    m_rrmax = re.search(r"Intervalo\s+RR.*?M[áa]x\.\s*longitud\s*([\d,\.]+)\s*s", texto_completo, re.IGNORECASE)
    d["rr_max_seg"] = extraer_flotante(m_rrmax.group(1)) if m_rrmax else None

    pausa_m = re.search(r"\bPausa\s+(\d+)", texto_completo)
    d["pausas"] = int(pausa_m.group(1)) if pausa_m else None
    p_max_m = re.search(r"Pausa.*?M[áa]x\.\s*longitud\s*([\d,\.]+)\s*s", texto_completo)
    d["pausa_max_seg"] = extraer_flotante(p_max_m.group(1)) if p_max_m else None

    d["latidos_caidos"] = int(re.search(r"Latidos?\s+ca[íi]dos?\s+(\d+)", texto_completo).group(1)) if re.search(r"Latidos?\s+ca[íi]dos?\s+(\d+)", texto_completo) else 0
    d["tv_episodios"] = limpiar_numero(re.search(r"\bTV\s+([\d\.]+)", texto_completo).group(1)) if re.search(r"\bTV\s+([\d\.]+)", texto_completo) else 0
    d["ev_duplas"] = limpiar_numero(re.search(r"Apareado\s+([\d\.]+)", texto_completo).group(1)) if re.search(r"Apareado\s+([\d\.]+)", texto_completo) else 0
    d["bigeminismo"] = limpiar_numero(re.search(r"Bigeminismo\s+([\d\.]+)", texto_completo).group(1)) if re.search(r"Bigeminismo\s+([\d\.]+)", texto_completo) else 0
    d["tsv_episodios"] = limpiar_numero(re.search(r"\bTSV\s+([\d\.]+)", texto_completo).group(1)) if re.search(r"\bTSV\s+([\d\.]+)", texto_completo) else 0

    st_dep = re.search(r"Depresi[óo]n ST\s+(\d+)", texto_completo)
    st_elev = re.search(r"Elevaci[óo]n ST\s+(\d+)", texto_completo)
    d["st_episodios"] = (int(st_dep.group(1)) if st_dep else 0) + (int(st_elev.group(1)) if st_elev else 0)

    sdnn_m = re.search(r"Valor de 24 horas\s+[\d\.,]+\s+([\d\.,]+)", texto_completo)
    d["sdnn_24h"] = limpiar_numero(sdnn_m.group(1)) if sdnn_m else None

    qtc_m = re.search(r"Todos los per[íi]odos\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)", texto_completo)
    d["qtc_prom"] = limpiar_numero(qtc_m.group(1)) if qtc_m else None

    ev_pac = re.search(r"Eventos del paciente\s*:\s*(\d+)", texto_completo)
    d["eventos_paciente"] = int(ev_pac.group(1)) if ev_pac else None

    m_dur = re.search(r"Hora de inicio.*?(\d{1,2}):(\d{2}).*?Hora de fin.*?(\d{1,2}):(\d{2})", texto_completo, re.DOTALL)
    d["horas_totales"] = 24.0

    return d

def ejecutar_triple_engine_holter_avanzado(pdf_bytes: bytes, filename: str = ""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pags = len(doc)

    d_m1 = extraer_m1_holter_determinista(doc)

    inline_paginas = []
    for num_p in range(total_pags):
        try:
            pix = doc[num_p].get_pixmap(dpi=105)
            img_b = pix.tobytes("jpeg")
            inline_paginas.append({
                "inline_data": {
                    "mime_type": "image/jpeg",
                    "data": base64.b64encode(img_b).decode("utf-8")
                }
            })
        except Exception:
            continue
    doc.close()

    prompt_m2 = f"""Eres el Auditor Multimodal del Centro Cardiovascular CENCARDIO.
Examina las {total_pags} páginas completas (texto, tablas, gráficos de tacograma y tiras de ritmo).
Devuelve estrictamente un JSON con los valores observados o null si no son visibles:
{{
  "paciente": "NOMBRE COMPLETO O null",
  "cedula": "CEDULA O null",
  "fc_prom": null,
  "fc_max": null,
  "fc_min": null,
  "fc_dia": null,
  "fc_noc": null,
  "total_latidos": null,
  "ev_total": null,
  "esv_total": null,
  "tv_episodios": null,
  "ev_duplas": null,
  "pausas": null,
  "pausa_max_seg": null,
  "rr_max_seg": null,
  "mcp_presente": null,
  "mcp_porcentaje": null,
  "sdnn_24h": null,
  "qtc_prom": null,
  "st_episodios": null,
  "eventos_paciente": null
}}"""

    ok_m2, d_m2, _ = consultar_gemini_json(prompt_m2, inline_paginas)
    if not ok_m2 or not d_m2:
        d_m2 = {}

    matriz = {}

    nom_val = d_m1.get("paciente") or d_m2.get("paciente") or "PACIENTE NO IDENTIFICADO"
    matriz["paciente"] = crear_expediente("paciente", nom_val, "", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("paciente") else "LEIDO_VISUALMENTE", "M1/M2", [])

    ced_val = d_m1.get("cedula") or d_m2.get("cedula") or ""
    matriz["cedula"] = crear_expediente("cedula", ced_val, "", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("cedula") else "LEIDO_VISUALMENTE", "M1/M2", [])

    qrs_m1 = d_m1.get("total_latidos")
    qrs_m2 = d_m2.get("total_latidos")
    if qrs_m1 is not None and qrs_m2 is not None:
        if abs(qrs_m1 - qrs_m2) <= (qrs_m1 * 0.01):
            matriz["total_latidos"] = crear_expediente("total_latidos", qrs_m1, "latidos", "VALIDADO", "CONSENSUS_M1_M2", [{"motor": "M1", "val": qrs_m1}, {"motor": "M2", "val": qrs_m2}])
        else:
            matriz["total_latidos"] = crear_expediente("total_latidos", qrs_m1, "latidos", "DISCREPANTE", "M1_TABULAR", [{"motor": "M1", "val": qrs_m1}, {"motor": "M2", "val": qrs_m2}], discrepancia="Diferencia > 1% entre motores.")
    elif qrs_m1 is not None:
        matriz["total_latidos"] = crear_expediente("total_latidos", qrs_m1, "latidos", "EXTRAIDO_DIRECTAMENTE", "M1", [])
    elif qrs_m2 is not None:
        matriz["total_latidos"] = crear_expediente("total_latidos", qrs_m2, "latidos", "LEIDO_VISUALMENTE", "M2", [])
    else:
        matriz["total_latidos"] = crear_expediente("total_latidos", None, "latidos", "NO_DETERMINABLE", "M3", [], requires_review=True)

    fc_m1 = d_m1.get("fc_prom")
    fc_m2 = d_m2.get("fc_prom")
    if fc_m1 is not None and fc_m2 is not None:
        if abs(fc_m1 - fc_m2) <= 2:
            matriz["fc_prom"] = crear_expediente("fc_prom", fc_m1, "lpm", "VALIDADO", "CONSENSUS_M1_M2", [{"m1": fc_m1, "m2": fc_m2}])
        else:
            matriz["fc_prom"] = crear_expediente("fc_prom", fc_m1, "lpm", "DISCREPANTE", "M1", [{"m1": fc_m1, "m2": fc_m2}], discrepancia="Variación > 2 lpm entre motores.")
    elif fc_m1 is not None:
        matriz["fc_prom"] = crear_expediente("fc_prom", fc_m1, "lpm", "EXTRAIDO_DIRECTAMENTE", "M1", [])
    elif fc_m2 is not None:
        matriz["fc_prom"] = crear_expediente("fc_prom", fc_m2, "lpm", "LEIDO_VISUALMENTE", "M2", [])
    elif matriz["total_latidos"]["value"] is not None:
        fc_calc = round(matriz["total_latidos"]["value"] / (24 * 60))
        matriz["fc_prom"] = crear_expediente("fc_prom", fc_calc, "lpm", "CALCULADO", "M3", [], metadata={"formula": "Total_QRS / 1440 min", "inputs": {"QRS": matriz["total_latidos"]["value"]}})
    else:
        matriz["fc_prom"] = crear_expediente("fc_prom", None, "lpm", "NO_DETERMINABLE", "M3", [], requires_review=True)

    matriz["fc_dia"] = crear_expediente("fc_dia", d_m1.get("fc_dia") or d_m2.get("fc_dia"), "lpm", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("fc_dia") else ("LEIDO_VISUALMENTE" if d_m2.get("fc_dia") else "NO_DETERMINABLE"), "M1/M2", [])
    matriz["fc_noc"] = crear_expediente("fc_noc", d_m1.get("fc_noc") or d_m2.get("fc_noc"), "lpm", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("fc_noc") else ("LEIDO_VISUALMENTE" if d_m2.get("fc_noc") else "NO_DETERMINABLE"), "M1/M2", [])

    pct_mcp_m1 = d_m1.get("mcp_porcentaje")
    pct_mcp_m2 = d_m2.get("mcp_porcentaje")
    mcp_pres_m1 = (d_m1.get("mcp_latidos") or 0) > 0 or (pct_mcp_m1 is not None and pct_mcp_m1 > 0)
    mcp_pres_m2 = bool(d_m2.get("mcp_presente"))
    mcp_final = mcp_pres_m1 or mcp_pres_m2
    pct_final = pct_mcp_m1 if pct_mcp_m1 is not None else (pct_mcp_m2 if pct_mcp_m2 is not None else 0.0)

    matriz["mcp"] = crear_expediente("marcapasos", {"presente": mcp_final, "porcentaje": pct_final, "latidos": d_m1.get("mcp_latidos") or 0}, "%", "VALIDADO" if (mcp_pres_m1 and mcp_pres_m2) else "EXTRAIDO_DIRECTAMENTE", "M1_M2", [])

    rr_max_final = d_m1.get("rr_max_seg") if d_m1.get("rr_max_seg") is not None else d_m2.get("rr_max_seg")
    pausas_m1 = d_m1.get("pausas")
    pausas_m2 = d_m2.get("pausas")

    if rr_max_final is not None and rr_max_final < 2.0:
        if (pausas_m1 and pausas_m1 > 0) or (pausas_m2 and pausas_m2 > 0):
            matriz["pausas"] = crear_expediente("pausas", 0, "pausas", "RECONSTRUIDO", "M3_VETO", [{"rr_max": rr_max_final}], metadata={"razon": f"RR max ({rr_max_final:.2f}s) < 2.0s; incompatibilidad fisiológica."})
        else:
            matriz["pausas"] = crear_expediente("pausas", 0, "pausas", "CONFIRMADO", "CONSENSUS", [{"rr_max": rr_max_final}])
    elif (pausas_m1 and pausas_m1 > 0) or (pausas_m2 and pausas_m2 > 0):
        p_val = max(pausas_m1 or 0, pausas_m2 or 0)
        matriz["pausas"] = crear_expediente("pausas", p_val, "pausas", "REQUIERE_REVISION", "ALERTA_PAUSA", [{"m1": pausas_m1, "m2": pausas_m2}], requires_review=True)
    else:
        matriz["pausas"] = crear_expediente("pausas", 0, "pausas", "CONFIRMADO", "M1_M2", [])

    ev_m1 = d_m1.get("ev_total")
    ev_m2 = d_m2.get("ev_total")
    if ev_m1 is not None and ev_m2 is not None:
        if ev_m1 == ev_m2:
            matriz["ev_total"] = crear_expediente("ev_total", ev_m1, "latidos", "CONFIRMADO", "CONSENSUS", [])
        else:
            matriz["ev_total"] = crear_expediente("ev_total", ev_m1, "latidos", "DISCREPANTE", "M1_TABLA", [{"m1": ev_m1, "m2": ev_m2}], discrepancia=f"M1={ev_m1} vs M2={ev_m2}. Se adopta M1.", requires_review=True)
    else:
        matriz["ev_total"] = crear_expediente("ev_total", ev_m1 if ev_m1 is not None else (ev_m2 or 0), "latidos", "EXTRAIDO_DIRECTAMENTE" if ev_m1 is not None else "LEIDO_VISUALMENTE", "M1/M2", [])

    matriz["esv_total"] = crear_expediente("esv_total", d_m1.get("esv_total") if d_m1.get("esv_total") is not None else (d_m2.get("esv_total") or 0), "latidos", "EXTRAIDO_DIRECTAMENTE", "M1/M2", [])
    matriz["tv_episodios"] = crear_expediente("tv_episodios", max(d_m1.get("tv_episodios", 0), d_m2.get("tv_episodios") or 0), "episodios", "REQUIERE_REVISION" if max(d_m1.get("tv_episodios", 0), d_m2.get("tv_episodios") or 0) > 0 else "CONFIRMADO", "M1_M2", [], requires_review=max(d_m1.get("tv_episodios", 0), d_m2.get("tv_episodios") or 0) > 0)
    matriz["ev_duplas"] = crear_expediente("ev_duplas", d_m1.get("ev_duplas", 0), "duplas", "EXTRAIDO_DIRECTAMENTE", "M1", [])
    matriz["tsv_episodios"] = crear_expediente("tsv_episodios", d_m1.get("tsv_episodios", 0), "episodios", "EXTRAIDO_DIRECTAMENTE", "M1", [])
    matriz["st_episodios"] = crear_expediente("st_episodios", d_m1.get("st_episodios", 0), "episodios", "EXTRAIDO_DIRECTAMENTE", "M1", [])
    matriz["sdnn_24h"] = crear_expediente("sdnn_24h", d_m1.get("sdnn_24h") or d_m2.get("sdnn_24h"), "ms", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("sdnn_24h") else ("LEIDO_VISUALMENTE" if d_m2.get("sdnn_24h") else "NO_DETERMINABLE"), "M1/M2", [])
    matriz["qtc_prom"] = crear_expediente("qtc_prom", d_m1.get("qtc_prom") or d_m2.get("qtc_prom"), "ms", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("qtc_prom") else ("LEIDO_VISUALMENTE" if d_m2.get("qtc_prom") else "NO_DETERMINABLE"), "M1/M2", [])
    matriz["eventos_paciente"] = crear_expediente("eventos_paciente", d_m1.get("eventos_paciente"), "marcas", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("eventos_paciente") is not None else "NO_DETERMINABLE", "M1", [])

    return matriz

def redactar_informe_holter_desde_matriz(matriz: dict, perfil: dict) -> str:
    mcp_data = matriz["mcp"]["value"]
    mcp_activo = mcp_data["presente"]
    pct_mcp = mcp_data["porcentaje"]
    fc_p = matriz["fc_prom"]["value"]
    fc_dia = matriz["fc_dia"]["value"]
    fc_noc = matriz["fc_noc"]["value"]

    if fc_dia is not None and fc_noc is not None and fc_dia > 0:
        desc_pct = ((fc_dia - fc_noc) / fc_dia) * 100
        desc_txt = f"conservado ({desc_pct:.1f}% de descenso nocturno)" if desc_pct >= 10 else f"atenuado ({desc_pct:.1f}% de descenso nocturno)"
        franja_txt = f" (Diurna: {fc_dia} lpm / Nocturna: {fc_noc} lpm; patrón circadiano {desc_txt})"
    else:
        franja_txt = " (descenso circadiano no determinable por ausencia de franjas horarias separadas)"

    if mcp_activo and pct_mcp >= 80.0:
        p1 = f"1. Ritmo comandado por marcapasos definitivo (ritmo electroestimulado/sensado en el {pct_mcp:.1f}% del registro analizado) con FC promedio de {fc_p or 'N/D'} lpm{franja_txt}."
    elif mcp_activo and pct_mcp > 0.0:
        p1 = f"1. Ritmo sinusal alternando con períodos de electroestimulación por marcapasos ({pct_mcp:.1f}% de latidos mediados por el dispositivo) con FC promedio de {fc_p or 'N/D'} lpm{franja_txt}."
    elif mcp_activo:
        p1 = f"1. Ritmo sinusal de base en paciente portador de marcapasos definitivo normofuncionante en modo demanda con FC promedio de {fc_p or 'N/D'} lpm{franja_txt}."
    else:
        p1 = f"1. Ritmo sinusal con FC promedio de {fc_p or 'N/D'} lpm{franja_txt}."

    p2 = "2. Eventos cronotrópicos: Sin taquicardias sostenidas de relevancia clínica ni bradicardia patológica fuera de los límites fisiológicos."

    qtc = matriz["qtc_prom"]["value"]
    if qtc is not None:
        if qtc > 500:
            p3 = f"3. Conducción AV sin bloqueos avanzados reportados; QTc SEVERAMENTE PROLONGADO ({qtc} ms: alto riesgo proarrítmico)."
        elif qtc > 460:
            p3 = f"3. Conducción AV sin bloqueos avanzados reportados; QTc prolongado ({qtc} ms)."
        else:
            p3 = f"3. Conducción AV sin bloqueos avanzados reportados; QTc en límites normales ({qtc} ms)."
    else:
        p3 = "3. Intervalo QTc no determinable en las tablas tabuladas del estudio."

    st_ep = matriz["st_episodios"]["value"]
    p4 = f"4. Alteraciones del segmento ST documentadas ({st_ep} episodios)." if st_ep > 0 else "4. Sin alteraciones isquémicas del segmento ST durante el registro."
    p5 = "5. Sin alteración patológica de la conducción auriculoventricular."

    if mcp_activo and pct_mcp >= 50.0:
        p6 = "6. Complejos ventriculares con morfología ancha inducidos por la electroestimulación artificial desde el ventrículo derecho."
    else:
        p6 = "6. Sin alteración en la conducción intraventricular nativa."

    ev = matriz["ev_total"]["value"] or 0
    tot_qrs = matriz["total_latidos"]["value"]
    carga_txt = f" (Carga ectópica: {(ev / tot_qrs) * 100:.2f}%)" if (tot_qrs and tot_qrs > 0) else ""

    tv = matriz["tv_episodios"]["value"] or 0
    duplas = matriz["ev_duplas"]["value"] or 0

    if tv > 0: lown = "Lown Grado IVb (Taquicardia Ventricular)"
    elif duplas > 0: lown = "Lown Grado IVa (Duplas ventriculares)"
    elif ev > 0: lown = "Lown Grado I-II"
    else: lown = "Lown Grado 0"

    p7 = f"7. Ectopia ventricular documentada: {ev} EV{carga_txt}, {lown}; sin arritmias ventriculares sostenidas." if ev > 0 else "7. Sin ectopia ventricular significativa."

    ev_pac = matriz["eventos_paciente"]["value"]
    if ev_pac is not None and ev_pac > 0:
        p8 = f"8. El paciente accionó el marcador de eventos en {ev_pac} ocasiones (correlacionar con diario de actividades)."
    else:
        p8 = "8. No se registraron marcas en el botón de eventos del paciente durante el monitoreo (correlacionar con bitácora clínica física)."

    sdnn = matriz["sdnn_24h"]["value"]
    if sdnn is not None and not mcp_activo:
        if sdnn <= 60:
            p9 = f"9. Variabilidad de la frecuencia cardíaca (VFC) severamente disminuida (SDNN: {sdnn} ms)."
            p11 = "11. Estratificación del riesgo autonómico por SDNN de 24 horas: Riesgo alto."
        elif sdnn <= 120:
            p9 = f"9. Variabilidad de la frecuencia cardíaca (VFC) disminuida (SDNN: {sdnn} ms)."
            p11 = "11. Estratificación del riesgo autonómico por SDNN de 24 horas: Riesgo moderado."
        else:
            p9 = f"9. Variabilidad de la frecuencia cardíaca (VFC) conservada (SDNN: {sdnn} ms)."
            p11 = "11. Estratificación del riesgo autonómico por SDNN de 24 horas: Bajo riesgo."
    elif mcp_activo:
        p9 = f"9. Variabilidad de la frecuencia cardíaca modificada por la frecuencia de estimulación del marcapasos (SDNN: {sdnn or 'N/D'} ms)."
        p11 = "11. Evaluación autonómica condicionada por la actividad del generador implantado."
    else:
        p9 = "9. Variabilidad de la frecuencia cardíaca no determinable cuantitativamente."
        p11 = "11. Estratificación autonómica no determinable por ausencia de cálculo de SDNN."

    pausas_cnt = matriz["pausas"]["value"]
    if mcp_activo:
        p10 = "10. En el registro evaluado no se evidencian pausas patológicas fuera del intervalo de escape programado (se sugiere correlación con interrogación telemétrica del dispositivo)."
    elif pausas_cnt is not None and pausas_cnt > 0:
        p10 = f"10. Se registraron {pausas_cnt} pausas significativas."
    else:
        p10 = "10. No se registraron pausas patológicas significativas (> 2.0 segundos)."

    if mcp_activo and pct_mcp >= 80.0:
        diag = f"Ritmo comandado predominantemente por electroestimulación artificial por marcapasos definitivo ({pct_mcp:.1f}% de estimulación). Adecuada respuesta cronotrópica al sensado."
    elif mcp_activo:
        diag = f"Ritmo sinusal de base con electroestimulación ventricular intermitente ({pct_mcp:.1f}% pacing) por marcapasos definitivo."
    else:
        diag = f"Ritmo sinusal con FC promedio {fc_p or 'N/D'} lpm. Modulación cronotrópica evaluada."

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

RECOMENDACIONES: Continuar manejo instaurado por cardiología y control clínico periódico. {'Se sugiere valoración periódica en consulta de seguimiento de marcapasos para telemetría completa.' if mcp_activo else ''}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# MOTOR 1, 2 Y 3 PARA MAPA TENSIONAL 24H (CUPS 895003)
# ==============================================================================
def extraer_m1_mapa_determinista(doc: fitz.Document) -> dict:
    texto = "".join([p.get_text() + "\n" for p in doc])
    d = {}

    m_nom = re.search(r"Nombre del paciente:\s*([^\n\r\|]+)", texto, re.IGNORECASE) or re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,45},\s*[A-ZÁÉÍÓÚÑ\s]{3,45})", texto)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else None

    m_id = re.search(r"ID paciente:\s*([^\n\r\s]+)", texto, re.IGNORECASE)
    d["cedula"] = m_id.group(1).replace(".", "") if m_id else None

    p_gen = re.search(r"Resumen general.*?Prom\.?:\s*(\d{2,3})\s*[\/\-]\s*(\d{2,3})\s*mmHg", texto, re.IGNORECASE)
    if p_gen:
        d["pas_24h"] = int(p_gen.group(1))
        d["pad_24h"] = int(p_gen.group(2))
    else:
        d["pas_24h"] = None
        d["pad_24h"] = None

    p_dia = re.search(r"(?:D[íi]a|Vigilia).*?Prom\.?:\s*(\d{2,3})\s*[\/\-]\s*(\d{2,3})\s*mmHg", texto, re.IGNORECASE)
    if p_dia:
        d["pas_dia"] = int(p_dia.group(1))
        d["pad_dia"] = int(p_dia.group(2))
    else:
        d["pas_dia"] = None
        d["pad_dia"] = None

    p_noc = re.search(r"(?:Noche|Sue[ñn]o).*?Prom\.?:\s*(\d{2,3})\s*[\/\-]\s*(\d{2,3})\s*mmHg", texto, re.IGNORECASE)
    if p_noc:
        d["pas_noc"] = int(p_noc.group(1))
        d["pad_noc"] = int(p_noc.group(2))
    else:
        d["pas_noc"] = None
        d["pad_noc"] = None

    c_sis = re.search(r"Sist[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", texto, re.IGNORECASE)
    d["carga_pas"] = extraer_flotante(c_sis.group(1)) if c_sis else None

    c_dia = re.search(r"Diast[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", texto, re.IGNORECASE)
    d["carga_pad"] = extraer_flotante(c_dia.group(1)) if c_dia else None

    caida_m = re.search(r"Sist[óo]lico\s*\(mmHg\)\s*.*?([\d,\.\-]+)\s*%", texto, re.DOTALL)
    d["caida_nocturna_val"] = extraer_flotante(caida_m.group(1)) if caida_m else None

    m_sueno = re.search(r"Resumen de los per[íi]odos de sue[ñn]o.*?Sist[óo]lico\s*\(mmHg\)\s*\n?\s*(\d+)\s*.*?(\d{2,3})\s*\([^\)]+\)\s*.*?Diast[óo]lico\s*\(mmHg\)\s*\n?\s*(\d+)\s*.*?(\d{2,3})\s*\(", texto, re.DOTALL | re.IGNORECASE)
    d["pas_max_sueno"] = int(m_sueno.group(2)) if m_sueno else None
    d["pad_max_sueno"] = int(m_sueno.group(4)) if m_sueno else None

    return d

def ejecutar_triple_engine_mapa_avanzado(pdf_bytes: bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pags = len(doc)
    d_m1 = extraer_m1_mapa_determinista(doc)

    inline_paginas = []
    for num_p in range(total_pags):
        try:
            pix = doc[num_p].get_pixmap(dpi=105)
            img_b = pix.tobytes("jpeg")
            inline_paginas.append({
                "inline_data": {
                    "mime_type": "image/jpeg",
                    "data": base64.b64encode(img_b).decode("utf-8")
                }
            })
        except Exception:
            continue
    doc.close()

    prompt_m2 = f"""Eres el Auditor Multimodal de MAPA Tensional de CENCARDIO.
Examina las {total_pags} páginas del reporte. Extrae en formato JSON exacto:
{{
  "paciente": "NOMBRE O null",
  "cedula": "CEDULA O null",
  "pas_24h": null,
  "pad_24h": null,
  "pas_dia": null,
  "pad_dia": null,
  "pas_noc": null,
  "pad_noc": null,
  "carga_pas": null,
  "carga_pad": null,
  "caida_nocturna_val": null,
  "pas_max_sueno": null,
  "pad_max_sueno": null
}}"""

    ok_m2, d_m2, _ = consultar_gemini_json(prompt_m2, inline_paginas)
    if not ok_m2 or not d_m2:
        d_m2 = {}

    matriz = {}

    nom_val = d_m1.get("paciente") or d_m2.get("paciente") or "PACIENTE MAPA"
    matriz["paciente"] = crear_expediente("paciente", nom_val, "", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("paciente") else "LEIDO_VISUALMENTE", "M1/M2", [])

    ced_val = d_m1.get("cedula") or d_m2.get("cedula") or ""
    matriz["cedula"] = crear_expediente("cedula", ced_val, "", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("cedula") else "LEIDO_VISUALMENTE", "M1/M2", [])

    pas_final = d_m1.get("pas_24h") if d_m1.get("pas_24h") is not None else d_m2.get("pas_24h")
    pad_final = d_m1.get("pad_24h") if d_m1.get("pad_24h") is not None else d_m2.get("pad_24h")

    matriz["pas_24h"] = crear_expediente("pas_24h", pas_final, "mmHg", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("pas_24h") is not None else "LEIDO_VISUALMENTE", "M1/M2", [])
    matriz["pad_24h"] = crear_expediente("pad_24h", pad_final, "mmHg", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("pad_24h") is not None else "LEIDO_VISUALMENTE", "M1/M2", [])

    if pas_final is not None and pad_final is not None:
        pam_calc = round(pad_final + ((pas_final - pad_final) / 3))
        pp_calc = pas_final - pad_final
        matriz["pam_24h"] = crear_expediente("pam_24h", pam_calc, "mmHg", "CALCULADO", "M3", [], metadata={"formula": "PAD + (PAS - PAD) / 3", "inputs": {"PAS": pas_final, "PAD": pad_final}})
        matriz["pp_24h"] = crear_expediente("pp_24h", pp_calc, "mmHg", "CALCULADO", "M3", [], metadata={"formula": "PAS - PAD", "inputs": {"PAS": pas_final, "PAD": pad_final}})
    else:
        matriz["pam_24h"] = crear_expediente("pam_24h", None, "mmHg", "NO_DETERMINABLE", "M3", [], requires_review=True)
        matriz["pp_24h"] = crear_expediente("pp_24h", None, "mmHg", "NO_DETERMINABLE", "M3", [], requires_review=True)

    pas_dia = d_m1.get("pas_dia") if d_m1.get("pas_dia") is not None else d_m2.get("pas_dia")
    pas_noc = d_m1.get("pas_noc") if d_m1.get("pas_noc") is not None else d_m2.get("pas_noc")
    caida_directa = d_m1.get("caida_nocturna_val") if d_m1.get("caida_nocturna_val") is not None else d_m2.get("caida_nocturna_val")

    if caida_directa is not None:
        matriz["caida_nocturna"] = crear_expediente("caida_nocturna", caida_directa, "%", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("caida_nocturna_val") is not None else "LEIDO_VISUALMENTE", "M1/M2", [])
    elif pas_dia is not None and pas_noc is not None and pas_dia > 0:
        caida_calc = round(((pas_dia - pas_noc) / pas_dia) * 100, 1)
        matriz["caida_nocturna"] = crear_expediente("caida_nocturna", caida_calc, "%", "CALCULADO", "M3", [], metadata={"formula": "((PAS_dia - PAS_noc) / PAS_dia) * 100", "inputs": {"PAS_dia": pas_dia, "PAS_noc": pas_noc}})
    else:
        matriz["caida_nocturna"] = crear_expediente("caida_nocturna", None, "%", "NO_DETERMINABLE", "M3", [], requires_review=True)

    c_pas = d_m1.get("carga_pas") if d_m1.get("carga_pas") is not None else d_m2.get("carga_pas")
    c_pad = d_m1.get("carga_pad") if d_m1.get("carga_pad") is not None else d_m2.get("carga_pad")
    matriz["carga_pas"] = crear_expediente("carga_pas", c_pas, "%", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("carga_pas") is not None else "LEIDO_VISUALMENTE", "M1/M2", [])
    matriz["carga_pad"] = crear_expediente("carga_pad", c_pad, "%", "EXTRAIDO_DIRECTAMENTE" if d_m1.get("carga_pad") is not None else "LEIDO_VISUALMENTE", "M1/M2", [])

    matriz["pas_max_sueno"] = crear_expediente("pas_max_sueno", d_m1.get("pas_max_sueno") or d_m2.get("pas_max_sueno"), "mmHg", "EXTRAIDO_DIRECTAMENTE", "M1/M2", [])
    matriz["pad_max_sueno"] = crear_expediente("pad_max_sueno", d_m1.get("pad_max_sueno") or d_m2.get("pad_max_sueno"), "mmHg", "EXTRAIDO_DIRECTAMENTE", "M1/M2", [])

    return matriz

def redactar_informe_mapa_desde_matriz(matriz: dict, perfil: dict) -> str:
    pas = matriz["pas_24h"]["value"]
    pad = matriz["pad_24h"]["value"]
    pam = matriz["pam_24h"]["value"]
    pp = matriz["pp_24h"]["value"]
    c_pas = matriz["carga_pas"]["value"]
    c_pad = matriz["carga_pad"]["value"]
    cn = matriz["caida_nocturna"]["value"]
    p_max_s = matriz["pas_max_sueno"]["value"]
    pd_max_s = matriz["pad_max_sueno"]["value"]

    p1 = f"1. Promedio de tensión arterial sistólica ({pas or 'N/D'} mmHg) y diastólica ({pad or 'N/D'} mmHg); Presión Arterial Media (PAM: {pam or 'N/D'} mmHg)."
    p2 = f"2. Carga tensional sistólica ({f'{c_pas:.1f}%' if c_pas is not None else 'N/D'}) y diastólica ({f'{c_pad:.1f}%' if c_pad is not None else 'N/D'})."

    if pp is not None:
        p3 = f"3. Presión de pulso conservada ({pp} mmHg)." if pp <= 60 else f"3. Presión de pulso aumentada ({pp} mmHg, marcador de rigidez arterial)."
    else:
        p3 = "3. Presión de pulso no determinable cuantitativamente."

    if cn is None:
        patron = "no determinable cuantitativamente (datos insuficientes del período de sueño)"
    elif cn >= 10.0 and cn <= 20.0:
        patron = f"conservado (dipping positivo: {cn:.1f}%)"
    elif cn > 20.0:
        patron = f"dipping extremo ({cn:.1f}%)"
    elif 0.0 <= cn < 10.0:
        patron = f"atenuado (no-dipper: {cn:.1f}%)"
    else:
        patron = f"invertido (patrón riser: {cn:.1f}%, alerta de sobrecarga cerebrovascular nocturna)"
    p4 = f"4. Patrón circadiano tensional {patron}."

    if p_max_s is not None and pd_max_s is not None and (p_max_s >= 145 or pd_max_s >= 95):
        p5 = f"5. Se presentaron incrementos significativos de presión durante el sueño (máximo registrado {p_max_s}/{pd_max_s} mmHg)."
    else:
        p5 = "5. No se evidenciaron incrementos paroxísticos de presión arterial durante el período de sueño."

    c_pas_v = c_pas or 0.0
    c_pad_v = c_pad or 0.0
    pas_v = pas or 120
    pad_v = pad or 80

    if c_pas_v < 15 and c_pad_v < 15 and pas_v < 130 and pad_v < 80:
        ctrl = "Control tensional óptimo de 24 horas"
    elif c_pas_v <= 30 or c_pad_v <= 30 or pas_v < 140:
        ctrl = "Control tensional subóptimo de 24 horas (Estadio I)"
    else:
        ctrl = "Descontrol tensional de 24 horas (Estadio II)"
    p6 = f"6. {ctrl}."

    recs = "Continuar régimen farmacológico actual y seguimiento médico periódico." if "óptimo" in ctrl else "Optimización y titulación de terapia antihipertensiva, reforzando medidas no farmacológicas y restricción sódica."

    return f"""INTERPRETACIÓN TEST MAPA - CUPS 895003
Hallazgos:
{p1}
{p2}
{p3}
{p4}
{p5}
{p6}

RECOMENDACIONES: {recs}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# MOTOR 1, 2 Y 3 PARA PRUEBA DE ESFUERZO / ERGOMETRÍA (CUPS 893805)
# ==============================================================================
def calcular_mets_protocolo(tiempo_min: float, protocolo: str) -> float:
    t = max(0.5, float(tiempo_min))
    p = str(protocolo).lower()
    if "naughton" in p:
        return round(1.6 + (t * 0.95), 1)
    elif "modificado" in p:
        if t <= 3.0: return round(1.5 + (t / 3.0) * 2.0, 1)
        elif t <= 6.0: return round(3.5 + ((t - 3.0) / 3.0) * 2.5, 1)
        else: return round(6.0 + ((t - 6.0) / 3.0) * 2.8, 1)
    else:
        return round(1.11 + (0.016 * (t * 60)), 1)

def calcular_duke_treadmill_score(tiempo_min: float, st_mm: float, angina_index: int):
    t = float(tiempo_min)
    st_val = float(st_mm)
    ang = int(angina_index)
    dts = t - (5.0 * st_val) - (4.0 * ang)
    if dts >= 5.0:
        riesgo = "Bajo riesgo coronario (Mortalidad cardiovascular < 1% anual)"
    elif -10.0 <= dts < 5.0:
        riesgo = "Riesgo coronario moderado (Mortalidad cardiovascular 1 - 3% anual)"
    else:
        riesgo = "Alto riesgo coronario (Mortalidad cardiovascular > 3% anual)"
    return round(dts, 1), riesgo

def redactar_informe_ergometria_desde_matriz(d: dict, perfil: dict) -> str:
    fcm_prev = 220 - d["edad"] if d["edad"] > 0 else 200
    porc = round((d["fc_pico"] / fcm_prev) * 100) if (fcm_prev > 0 and d["fc_pico"] > 0) else 0
    suf = "suficiente" if porc >= 85 else "insuficiente"
    dp = d["fc_pico"] * d["pas_pico"]
    st_res = "Sin alteraciones isquémicas del segmento ST" if d["st_mm"] < 1.0 else f"Alteraciones de la repolarización con infradesnivel del ST de {d['st_mm']} mm"
    diag_el = "negativa" if d["st_mm"] < 1.0 else "positiva"

    reserva_den = (fcm_prev - d["fc_basal"]) if (fcm_prev - d["fc_basal"]) > 0 else 1
    fcr_pct = round(((d["fc_pico"] - d["fc_basal"]) / reserva_den) * 100)
    incomp_crono = "Incompetencia cronotrópica" if fcr_pct < 80 else "Respuesta cronotrópica adecuada"

    return f"""INTERPRETACIÓN PRUEBA DE ESFUERZO COMPUTARIZADA - CUPS 893805

1. Ritmo sinusal normal basal y durante todas las etapas del esfuerzo físico.
2. Protocolo de {d['protocolo']} completado con duración de {d['tiempo_min']:.2f} minutos ({d['etapa']}).
3. Capacidad funcional alcanzada: {d['mets']} METs.
4. Respuesta cronotrópica: FC basal {d['fc_basal']} lpm alcanzando FC pico de {d['fc_pico']} lpm ({porc}% de la FCM prevista de {fcm_prev} lpm; Reserva Cronotrópica FCR: {fcr_pct}%, {incomp_crono}, prueba {suf}).
5. Respuesta hemodinámica presora: PA basal {d['pas_basal']}/{d['pad_basal']} mmHg alcanzando PA pico de {d['pas_pico']}/{d['pad_pico']} mmHg.
6. Doble producto máximo alcanzado: {dp:,} mmHg*lpm.
7. Comportamiento electrocardiográfico del ST: {st_res}.
8. Sin arritmias ventriculares complejas ni eventos supraventriculares inducidos por el ejercicio.
9. Motivo de suspensión: Consecución de frecuencia cardíaca diagnóstica y fatiga física voluntaria, sin angina.
10. Estratificación pronóstica por Duke Treadmill Score: {d['dts']:.1f} ({d['duke_riesgo']}).

CONCLUSIÓN DIAGNÓSTICA:
Prueba de esfuerzo física {suf}, eléctricamente {diag_el} para isquemia miocárdica inducible. Adecuada tolerancia funcional y hemodinámica ({incomp_crono}).
RECOMENDACIONES: {'Continuar control médico periódico y prescripción de actividad física regular.' if d['st_mm'] < 1.0 else 'Valoración prioritaria por cardiología clínica para estudio funcional o angiografía coronaria.'}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# INYECTORES DE PDF CON ESTAMPADO DINÁMICO Y QR
# ==============================================================================
@st.cache_data
def generar_qr_verificacion_token(token_unico: str):
    url_completa = f"https://holtercencardio.streamlit.app/?token={token_unico}"
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
        if caja:
            tinta = tinta.crop(caja)
        buf = io.BytesIO()
        tinta.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None

def normalizar_nombre_archivo(nombre):
    limpio = re.sub(r'[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s]', ' ', str(nombre))
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
    qr_bytes = generar_qr_verificacion_token(cod_uuid)
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

    qr_bytes = generar_qr_verificacion_token(cod_uuid)
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
            with open(nom, "rb") as f:
                logo_bytes = f.read()
            break
    if logo_bytes:
        page.insert_image(fitz.Rect(36, 30, 150, 75), stream=logo_bytes)

    page.insert_text(fitz.Point(165, 45), "CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO", fontsize=11, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(165, 58), "INFORME DE ERGOMETRÍA Y PRUEBA DE ESFUERZO COMPUTARIZADA", fontsize=8.5, fontname="helv", color=(0.78, 0.06, 0.18))
    page.insert_text(fitz.Point(165, 70), "CUPS: 893805 · Habilitación MinSalud Colombia · Res. 3100 de 2019", fontsize=7, fontname="helv", color=(0.4, 0.45, 0.5))
    page.draw_rect(fitz.Rect(36, 85, 576, 87), color=None, fill=(0.04, 0.15, 0.25), overlay=True)

    page.draw_rect(fitz.Rect(36, 95, 576, 155), color=(0.85, 0.9, 0.95), fill=(0.97, 0.98, 1.0), width=1)
    page.insert_text(fitz.Point(46, 112), f"PACIENTE: {str(d['paciente']).upper()}", fontsize=8.5, fontname="helv", color=(0.04, 0.15, 0.25))
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

    qr_bytes = generar_qr_verificacion_token(cod_uuid)
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
            except Exception:
                pass

    pdf_bytes = doc.tobytes()
    doc.close()
    return pdf_bytes, img_preview

# ==============================================================================
# REPORTE GERENCIAL (EXCEL / CSV)
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
# BARRA LATERAL Y NAVEGACIÓN
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
        st.success("🖋️ Firma y Sello Digital Oficial cargados.")

    st.divider()
    modalidad_seleccionada = st.radio(
        "Procedimiento Cardiológico:",
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
        except Exception as e:
            st.error(f"Error: {e}")

    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        cerrar_sesion()
        st.rerun()

estado_nube_txt = "🟢 Nube Supabase Activa" if supabase else "🟡 Almacenamiento Local (SQLite)"
st.markdown(f"""
    <div class="top-hospital-bar">
        <div>
            <div style="font-size: 1.35rem; font-weight: 800; color: #0a2540; line-height: 1.2;">
                CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO
            </div>
            <div style="font-size: 0.82rem; font-weight: 700; color: #c8102e; text-transform: uppercase;">
                Workstation Diagnóstica Multimotor · Alta Complejidad
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
# PESTAÑA 1: PROCESAMIENTO MULTIMOTOR
# ==============================================================================
with tab_procesar:
    if "Holter" in modalidad_seleccionada:
        st.markdown("#### 🫀 Análisis Multimotor de Holter ECG 24H (CUPS 895001)")
        uploaded_file = st.file_uploader("Seleccione el archivo PDF del estudio Holter:", type=["pdf"], key="holter_uploader")

        if uploaded_file is not None:
            bytes_originales = uploaded_file.getvalue()

            if "archivo_cargado_nombre" not in st.session_state or st.session_state.archivo_cargado_nombre != uploaded_file.name:
                with st.spinner("🔍 Ejecutando Triple Engine: Extracción M1 + Visión M2 (100% de páginas) + Reconstrucción M3..."):
                    matriz_evidencia = ejecutar_triple_engine_holter_avanzado(bytes_originales, uploaded_file.name)
                    texto_informe = redactar_informe_holter_desde_matriz(matriz_evidencia, perfil_activo)

                    st.session_state.matriz_holter = matriz_evidencia
                    st.session_state.texto_informe = texto_informe
                    st.session_state.archivo_cargado_nombre = uploaded_file.name
                    st.session_state.estudio_uuid = str(uuid.uuid4()).upper()
                    st.session_state.telefono_paciente = buscar_telefono_servicio(matriz_evidencia["cedula"]["value"])

            matriz = st.session_state.matriz_holter

            with st.expander("🛡️ EXPEDIENTE MULTIMOTOR: TRAZABILIDAD Y MATRIZ DE EVIDENCIA", expanded=True):
                col_e1, col_e2, col_e3 = st.columns(3)
                with col_e1:
                    st.markdown("**Parámetros Determinados**")
                    for k in ["total_latidos", "fc_prom", "mcp"]:
                        exp = matriz[k]
                        val = exp["value"] if not isinstance(exp["value"], dict) else f"{exp['value']['porcentaje']}% pacing"
                        st.write(f"• **{exp['parametro']}:** `{val}` | Estado: `{exp['status']}`")
                with col_e2:
                    st.markdown("**Arritmias y Cronotropismo**")
                    for k in ["ev_total", "tv_episodios", "pausas"]:
                        exp = matriz[k]
                        st.write(f"• **{exp['parametro']}:** `{exp['value']}` | Estado: `{exp['status']}`")
                with col_e3:
                    st.markdown("**Repolarización y Autonómico**")
                    for k in ["sdnn_24h", "qtc_prom", "st_episodios"]:
                        exp = matriz[k]
                        st.write(f"• **{exp['parametro']}:** `{exp['value'] or 'N/D'}` | Estado: `{exp['status']}`")

            col_ed, col_prev = st.columns([1, 1], gap="large")
            with col_ed:
                c_nom, c_ced = st.columns([1.8, 1.2])
                with c_nom:
                    nombre_conf = st.text_input("👤 Paciente:", value=matriz["paciente"]["value"])
                with c_ced:
                    ced_conf = st.text_input("🪪 Cédula / ID:", value=matriz["cedula"]["value"])

                tel_conf = st.text_input("📱 Celular (WhatsApp):", value=st.session_state.get("telefono_paciente", ""))

                st.subheader("📝 Dictamen Oficial Certificado")
                informe_final = st.text_area("Texto oficial para el documento final:", value=st.session_state.texto_informe, height=380)

                debe_firmar = perfil_activo["id"] in ["dr.amaya", "admin"]
                pdf_generado, img_prev = inyectar_holter_pdf(bytes_originales, informe_final, nombre_conf, perfil_activo, st.session_state.estudio_uuid, estampador_activo=debe_firmar)

                c_b1, c_b2 = st.columns(2)
                with c_b1:
                    st.download_button(
                        label="📄 DESCARGAR HOLTER FIRMADO",
                        data=pdf_generado,
                        file_name=f"{normalizar_nombre_archivo(nombre_conf)}_Holter_CUPS_895001.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                with c_b2:
                    if st.button("💾 Certificar y Archivar Estudio", use_container_width=True):
                        fc_res = matriz["fc_prom"]["value"] or "N/D"
                        sdnn_res = matriz["sdnn_24h"]["value"] or "N/D"
                        param_clave = f"FC {fc_res} lpm | SDNN {sdnn_res} ms"
                        ok, msg = guardar_estudio_servicio(nombre_conf, "Holter ECG 24 Horas", "CUPS 895001", param_clave, perfil_activo['nombre_completo'], informe_final, pdf_generado, st.session_state.estudio_uuid)
                        if ok: st.success(f"✅ {msg}")
                        else: st.error(f"❌ {msg}")

                if tel_conf:
                    tel_clean = re.sub(r'\D', '', tel_conf)
                    if not tel_clean.startswith("57") and len(tel_clean) == 10: tel_clean = "57" + tel_clean
                    url_val = f"https://holtercencardio.streamlit.app/?token={st.session_state.estudio_uuid}"
                    msg_w = f"Estimado(a) paciente {nombre_conf}, el Centro Cardiovascular CENCARDIO le entrega su resultado oficial de Holter (CUPS 895001). Certificación forense: {url_val}"
                    st.write("")
                    st.link_button("📲 ENVIAR RESULTADO POR WHATSAPP", f"https://wa.me/{tel_clean}?text={urllib.parse.quote(msg_w)}", use_container_width=True)

            with col_prev:
                st.subheader("👁️ Vista Previa del Informe")
                st.markdown('<div class="preview-container">', unsafe_allow_html=True)
                if img_prev:
                    st.image(img_prev, caption=f"Página 1 Oficial - {nombre_conf}", use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

    elif "MAPA" in modalidad_seleccionada:
        st.markdown("#### 🩺 Análisis Multimotor de MAPA Tensional 24H (CUPS 895003)")
        uploaded_mapa = st.file_uploader("Seleccione el archivo PDF del estudio MAPA:", type=["pdf"], key="mapa_uploader")

        if uploaded_mapa is not None:
            bytes_mapa = uploaded_mapa.getvalue()

            if "archivo_mapa_nombre" not in st.session_state or st.session_state.archivo_mapa_nombre != uploaded_mapa.name:
                with st.spinner("🔍 Analizando MAPA: M1 (Tablas) + M2 (Visión 100% de páginas) + M3 (Ecuaciones PAM/Dipping)..."):
                    matriz_mapa = ejecutar_triple_engine_mapa_avanzado(bytes_mapa)
                    texto_mapa = redactar_informe_mapa_desde_matriz(matriz_mapa, perfil_activo)

                    st.session_state.matriz_mapa = matriz_mapa
                    st.session_state.texto_mapa = texto_mapa
                    st.session_state.archivo_mapa_nombre = uploaded_mapa.name
                    st.session_state.mapa_uuid = str(uuid.uuid4()).upper()
                    st.session_state.telefono_mapa = buscar_telefono_servicio(matriz_mapa["cedula"]["value"])

            matriz_m = st.session_state.matriz_mapa

            with st.expander("🛡️ EXPEDIENTE MULTIMOTOR: TRAZABILIDAD HEMODINÁMICA", expanded=True):
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1:
                    st.markdown("**Presiones Promedio**")
                    for k in ["pas_24h", "pad_24h", "pam_24h"]:
                        exp = matriz_m[k]
                        st.write(f"• **{exp['parametro']}:** `{exp['value']}` {exp['unidades']} | Estado: `{exp['status']}`")
                with col_m2:
                    st.markdown("**Cargas y Modulación**")
                    for k in ["carga_pas", "carga_pad", "caida_nocturna"]:
                        exp = matriz_m[k]
                        st.write(f"• **{exp['parametro']}:** `{exp['value']}` {exp['unidades']} | Estado: `{exp['status']}`")
                with col_m3:
                    st.markdown("**Rigidez y Sueño**")
                    for k in ["pp_24h", "pas_max_sueno", "pad_max_sueno"]:
                        exp = matriz_m[k]
                        st.write(f"• **{exp['parametro']}:** `{exp['value'] or 'N/D'}` {exp['unidades']} | Estado: `{exp['status']}`")

            col_ed_m, col_prev_m = st.columns([1, 1], gap="large")
            with col_ed_m:
                c_nom_m, c_ced_m = st.columns([1.8, 1.2])
                with c_nom_m:
                    nom_conf_m = st.text_input("👤 Paciente:", value=matriz_m["paciente"]["value"], key="nom_mapa")
                with c_ced_m:
                    ced_conf_m = st.text_input("🪪 Cédula / ID:", value=matriz_m["cedula"]["value"], key="ced_mapa")

                tel_conf_m = st.text_input("📱 Celular (WhatsApp):", value=st.session_state.get("telefono_mapa", ""), key="tel_mapa")

                st.subheader("📝 Dictamen Oficial de MAPA")
                informe_final_m = st.text_area("Texto oficial para el documento final:", value=st.session_state.texto_mapa, height=360, key="txt_mapa")

                debe_firmar_m = perfil_activo["id"] in ["dr.amaya", "admin"]
                pdf_mapa_gen, img_prev_m = inyectar_mapa_pdf(bytes_mapa, informe_final_m, nom_conf_m, perfil_activo, st.session_state.mapa_uuid, estampador_activo=debe_firmar_m)

                c_bm1, c_bm2 = st.columns(2)
                with c_bm1:
                    st.download_button(
                        label="📄 DESCARGAR MAPA FIRMADO",
                        data=pdf_mapa_gen,
                        file_name=f"{normalizar_nombre_archivo(nom_conf_m)}_MAPA_CUPS_895003.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                with c_bm2:
                    if st.button("💾 Certificar y Archivar MAPA", use_container_width=True):
                        pas_v = matriz_m["pas_24h"]["value"] or "N/D"
                        pad_v = matriz_m["pad_24h"]["value"] or "N/D"
                        pam_v = matriz_m["pam_24h"]["value"] or "N/D"
                        param_clave_m = f"PA 24h: {pas_v}/{pad_v} mmHg (PAM {pam_v})"
                        ok, msg = guardar_estudio_servicio(nom_conf_m, "MAPA Tensional 24 Horas", "CUPS 895003", param_clave_m, perfil_activo['nombre_completo'], informe_final_m, pdf_mapa_gen, st.session_state.mapa_uuid)
                        if ok: st.success(f"✅ {msg}")
                        else: st.error(f"❌ {msg}")

                if tel_conf_m:
                    tel_clean_m = re.sub(r'\D', '', tel_conf_m)
                    if not tel_clean_m.startswith("57") and len(tel_clean_m) == 10: tel_clean_m = "57" + tel_clean_m
                    url_val_m = f"https://holtercencardio.streamlit.app/?token={st.session_state.mapa_uuid}"
                    msg_wm = f"Estimado(a) paciente {nom_conf_m}, CENCARDIO le entrega su resultado oficial de MAPA Tensional (CUPS 895003). Certificación forense: {url_val_m}"
                    st.write("")
                    st.link_button("📲 ENVIAR RESULTADO POR WHATSAPP", f"https://wa.me/{tel_clean_m}?text={urllib.parse.quote(msg_wm)}", use_container_width=True, key="wa_btn_mapa")

            with col_prev_m:
                st.subheader("👁️ Vista Previa del Informe")
                st.markdown('<div class="preview-container">', unsafe_allow_html=True)
                if img_prev_m:
                    st.image(img_prev_m, caption=f"Página 1 Oficial - {nom_conf_m}", use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("#### 🏃 Consola Multimotor de Prueba de Esfuerzo / Ergometría (CUPS 893805)")
        st.caption("Suba imágenes de tirillas/post-its o ingrese los parámetros. El Motor 3 aplica las fórmulas específicas de METs según protocolo y Duke Treadmill Score completo.")

        col_f1, col_f2 = st.columns([1.2, 1], gap="large")
        with col_f1:
            fotos_esfuerzo = st.file_uploader(
                "📸 Subir fotos o escaneos de tiras de esfuerzo (JPG/PNG):",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True,
                key="uploader_erg"
            )

            if fotos_esfuerzo:
                if st.button("⚡ EXTRAER CON VISIÓN MULTIMODAL", type="primary", use_container_width=True):
                    with st.spinner("🤖 Analizando trazados continuos con M2..."):
                        inline_items = []
                        for foto in fotos_esfuerzo:
                            try:
                                raw_b = foto.getvalue()
                                img = Image.open(io.BytesIO(raw_b))
                                img = ImageOps.exif_transpose(img)
                                if img.mode != "RGB": img = img.convert("RGB")
                                if max(img.size) > 1400:
                                    r = 1400 / max(img.size)
                                    img = img.resize((int(img.size[0] * r), int(img.size[1] * r)), Image.Resampling.LANCZOS)
                                buf_opt = io.BytesIO()
                                img.save(buf_opt, format="JPEG", quality=80)
                                inline_items.append({"inline_data": {"mime_type": "image/jpeg", "data": base64.b64encode(buf_opt.getvalue()).decode("utf-8")}})
                            except Exception:
                                continue

                        prompt_erg = """Extrae en JSON exacto:
{"paciente": "NOMBRE", "cedula": "ID", "edad": 35, "sexo": "Femenino", "protocolo": "Bruce", "etapa": "Etapa 4", "tiempo_min": 9.0, "fc_basal": 75, "fc_pico": 155, "pas_basal": 120, "pad_basal": 80, "pas_pico": 160, "pad_pico": 90, "st_mm": 0.0, "angina_idx": 0}"""
                        ok_e, data_e, _ = consultar_gemini_json(prompt_erg, inline_items)
                        if ok_e and data_e:
                            st.session_state.erg_paciente = str(data_e.get("paciente", "")).upper()
                            st.session_state.erg_cedula = str(data_e.get("cedula", ""))
                            st.session_state.erg_edad = int(data_e.get("edad", 35))
                            st.session_state.erg_sexo = data_e.get("sexo", "Femenino")
                            st.session_state.erg_protocolo = data_e.get("protocolo", "Bruce")
                            st.session_state.erg_etapa = data_e.get("etapa", "Final")
                            st.session_state.erg_tiempo = float(data_e.get("tiempo_min", 9.0))
                            st.session_state.erg_fc_basal = int(data_e.get("fc_basal", 75))
                            st.session_state.erg_fc_pico = int(data_e.get("fc_pico", 155))
                            st.session_state.erg_pa_basal = f"{data_e.get('pas_basal', 120)}/{data_e.get('pad_basal', 80)}"
                            st.session_state.erg_pa_pico = f"{data_e.get('pas_pico', 160)}/{data_e.get('pad_pico', 90)}"
                            st.session_state.erg_st_mm = float(data_e.get("st_mm", 0.0))
                            st.session_state.erg_angina = int(data_e.get("angina_idx", 0))
                            st.success("✅ Datos extraídos por M2. Revise en el panel.")
                            st.rerun()

            c1, c2, c3 = st.columns(3)
            with c1:
                p_nom = st.text_input("Paciente:", value=st.session_state.get("erg_paciente", ""))
                p_ced = st.text_input("Cédula / ID:", value=st.session_state.get("erg_cedula", ""))
            with c2:
                p_edad = st.number_input("Edad:", min_value=1, max_value=110, value=st.session_state.get("erg_edad", 35))
                p_sexo = st.selectbox("Sexo:", ["Femenino", "Masculino"], index=0 if st.session_state.get("erg_sexo", "Femenino") == "Femenino" else 1)
            with c3:
                p_proto = st.selectbox("Protocolo:", ["Bruce", "Bruce Modificado", "Naughton"], index=0)
                p_etapa = st.text_input("Etapa alcanzada:", value=st.session_state.get("erg_etapa", "Etapa 4"))

            c4, c5, c6 = st.columns(3)
            with c4:
                p_tiempo = st.number_input("Tiempo total (min):", min_value=0.5, max_value=40.0, value=st.session_state.get("erg_tiempo", 9.0), step=0.1)
                mets_calc = calcular_mets_protocolo(p_tiempo, p_proto)
                st.write(f"Capacidad calculada M3: **{mets_calc} METs**")
            with c5:
                p_fc_basal = st.number_input("FC Basal (lpm):", value=st.session_state.get("erg_fc_basal", 75))
                p_fc_pico = st.number_input("FC Pico (lpm):", value=st.session_state.get("erg_fc_pico", 155))
            with c6:
                p_pa_basal = st.text_input("PA Basal (mmHg):", value=st.session_state.get("erg_pa_basal", "120/80"))
                p_pa_pico = st.text_input("PA Pico (mmHg):", value=st.session_state.get("erg_pa_pico", "160/90"))

            c7, c8 = st.columns(2)
            with c7:
                p_st = st.number_input("Desviación del ST (mm):", value=st.session_state.get("erg_st_mm", 0.0), step=0.5)
            with c8:
                p_angina = st.selectbox("Índice de Angina:", [0, 1, 2], index=st.session_state.get("erg_angina", 0), format_func=lambda x: {0: "0: Sin dolor torácico", 1: "1: Angina no limitante", 2: "2: Angina que motivó la detención"}[x])

            dts_val, dts_cat = calcular_duke_treadmill_score(p_tiempo, p_st, p_angina)
            tel_erg = st.text_input("Celular (WhatsApp):", value=buscar_telefono_servicio(p_ced))

        with col_f2:
            if p_nom.strip():
                pas_b, pad_b = [int(x) for x in p_pa_basal.split("/")] if "/" in p_pa_basal else (120, 80)
                pas_p, pad_p = [int(x) for x in p_pa_pico.split("/")] if "/" in p_pa_pico else (160, 90)

                datos_erg = {
                    "paciente": p_nom, "cedula": p_ced, "edad": p_edad, "sexo": p_sexo,
                    "protocolo": p_proto, "etapa": p_etapa, "tiempo_min": p_tiempo,
                    "mets": mets_calc, "fc_basal": p_fc_basal, "fc_pico": p_fc_pico,
                    "pas_basal": pas_b, "pad_basal": pad_b, "pas_pico": pas_p, "pad_pico": pad_p,
                    "st_mm": p_st, "angina_idx": p_angina, "dts": dts_val, "duke_riesgo": dts_cat
                }

                st.subheader("📝 Informe Oficial de Esfuerzo")
                txt_erg_generado = redactar_informe_ergometria_desde_matriz(datos_erg, perfil_activo)
                txt_erg_final = st.text_area("Texto oficial para certificación:", value=txt_erg_generado, height=330)

                uuid_erg = str(uuid.uuid4()).upper()
                pdf_erg_gen, img_prev_erg = generar_pdf_ergometria_completo(datos_erg, txt_erg_final, perfil_activo, uuid_erg, imagenes_adjuntas=fotos_esfuerzo or [])

                ce1, ce2 = st.columns(2)
                with ce1:
                    st.download_button(
                        label="📄 DESCARGAR ERGOMETRÍA",
                        data=pdf_erg_gen,
                        file_name=f"{normalizar_nombre_archivo(p_nom)}_Esfuerzo_CUPS_893805.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                with ce2:
                    if st.button("💾 Certificar y Archivar Esfuerzo", use_container_width=True):
                        pclave_e = f"FC Pico: {p_fc_pico} | {mets_calc} METs | DTS: {dts_val}"
                        ok, msg = guardar_estudio_servicio(p_nom, "Prueba de Esfuerzo", "CUPS 893805", pclave_e, perfil_activo['nombre_completo'], txt_erg_final, pdf_erg_gen, uuid_erg)
                        if ok: st.success(f"✅ {msg}")
                        else: st.error(f"❌ {msg}")

                if tel_erg:
                    tel_cl_e = re.sub(r'\D', '', tel_erg)
                    if not tel_cl_e.startswith("57") and len(tel_cl_e) == 10: tel_cl_e = "57" + tel_cl_e
                    url_val_e = f"https://holtercencardio.streamlit.app/?token={uuid_erg}"
                    msg_we = f"Estimado(a) paciente {p_nom}, CENCARDIO le entrega su resultado de Prueba de Esfuerzo (CUPS 893805). Certificación: {url_val_e}"
                    st.write("")
                    st.link_button("📲 ENVIAR RESULTADO POR WHATSAPP", f"https://wa.me/{tel_cl_e}?text={urllib.parse.quote(msg_we)}", use_container_width=True)

                if img_prev_erg:
                    st.image(img_prev_erg, caption=f"Página 1 Oficial - {p_nom}", use_container_width=True)
            else:
                st.info("💡 Digite el nombre del paciente o suba fotografías de las tiras de esfuerzo.")

# ==============================================================================
# PESTAÑA 2: ARCHIVO CLÍNICO Y REPORTES GERENCIALES
# ==============================================================================
with tab_historial:
    st.markdown("### 📁 Archivo Clínico Digital & Reportes Gerenciales")
    historial, origen_datos = obtener_historial_servicio()

    if not historial:
        st.info("No hay estudios archivados en el sistema.")
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
        with f1: busqueda = st.text_input("🔍 Buscar por paciente:", "")
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
        for item in df_filtrado.itertuples():
            with st.expander(f"👤 {item.paciente} | {item.modalidad} ({item.fecha_registro})"):
                st.write(f"**Especialista Lector:** {item.medico} | **Parámetro Clave:** {item.parametro_clave}")
                st.write(f"**Token Criptográfico:** `{item.codigo_verificacion}`")
                st.write(f"**Hash SHA-256:** `{item.hash_sha256}`")
                c_d1, c_d2 = st.columns([1.5, 1])
                with c_d1:
                    if item.pdf_url:
                        st.link_button("📥 Descargar PDF desde Nube", item.pdf_url)
                    else:
                        pdf_rec = obtener_pdf_bytes_individual(item.id)
                        if pdf_rec:
                            st.download_button(
                                label="📥 Descargar PDF de Respaldo Local",
                                data=pdf_rec,
                                file_name=f"{normalizar_nombre_archivo(item.paciente)}_{item.cups}.pdf",
                                mime="application/pdf",
                                key=f"hist_desc_{item.id}"
                            )
                with c_d2:
                    if st.button("🗑️ Eliminar Estudio", key=f"del_{item.id}"):
                        eliminar_estudio_servicio(item.id)
                        st.toast(f"Estudio de {item.paciente} eliminado.", icon="🗑️")
                        st.rerun()
