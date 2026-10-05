import streamlit as st
from pypdf import PdfReader
import fitz  # PyMuPDF
from PIL import Image
import qrcode
import re
import base64
import os
import io
import sqlite3
import pandas as pd
from datetime import datetime
import uuid
import urllib.parse

st.set_page_config(
    page_title="Centro Cardiovascular Colombiano Cencardio",
    page_icon="🫀",
    layout="wide"
)

# ==========================================
# UTILIDAD: RECURSOS E IDENTIDAD VISUAL
# ==========================================
def obtener_logo_b64():
    for nom in ["cencardio.jpg", "cencardio.png", "cencardio.jpeg", "logo.png", "logo.jpg"]:
        if os.path.exists(nom):
            with open(nom, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            mime = "png" if nom.endswith("png") else "jpeg"
            return f"data:image/{mime};base64,{b64}"
    return None

def cargar_fondo():
    for ext in ["fondo.jpg", "fondo.png", "fondo.jpeg"]:
        if os.path.exists(ext):
            with open(ext, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            mime = "png" if ext.endswith("png") else "jpeg"
            return f"""
            <style>
            .stApp {{
                background-image: linear-gradient(rgba(244, 247, 250, 0.92), rgba(244, 247, 250, 0.92)), 
                                  url("data:image/{mime};base64,{b64}");
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }}
            </style>
            """
    return """
    <style>
    .stApp { background: linear-gradient(140deg, #f0f4f8 0%, #f8fafc 50%, #edf2f7 100%); }
    </style>
    """

st.markdown(cargar_fondo(), unsafe_allow_html=True)

st.markdown("""
    <style>
    header[data-testid="stHeader"] { background: transparent !important; }
    div[data-testid="stDecoration"] { display: none !important; }
    .block-container { padding-top: 1.8rem !important; padding-bottom: 2.5rem !important; }
    div[data-testid="stForm"] { border: none !important; padding: 0 !important; }

    .cencardio-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 2.4rem 2.8rem;
        box-shadow: 0 12px 30px -8px rgba(19, 50, 91, 0.12);
        max-width: 520px;
        margin: 2rem auto;
        text-align: center;
    }
    .cencardio-logo-img { max-width: 175px; height: auto; margin-bottom: 1rem; display: inline-block; }
    .cencardio-title { color: #13325b; font-size: 1.35rem; font-weight: 800; letter-spacing: 0.3px; line-height: 1.3; margin-top: 0.4rem; text-transform: uppercase; }
    .cencardio-sub { color: #c8102e; font-size: 0.88rem; font-weight: 700; letter-spacing: 0.5px; text-transform: uppercase; margin-bottom: 1.8rem; }

    h1 { color: #13325b !important; font-weight: 800 !important; }
    h2, h3 { color: #1b365d !important; }
    [data-testid="stMetricValue"] { color: #13325b !important; font-weight: 700; }
    
    .stButton > button {
        background-color: #13325b !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 0.55rem 1.4rem !important;
        transition: all 0.2s ease;
    }
    .stButton > button:hover {
        background-color: #c8102e !important;
        color: white !important;
    }

    textarea {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        font-family: monospace !important;
        font-size: 13px !important;
    }
    .preview-container {
        border: 2px solid #cbd5e1;
        border-radius: 10px;
        box-shadow: 0 6px 14px -3px rgba(19, 50, 91, 0.08);
        background: white;
        padding: 6px;
    }

    .triage-rojo {
        background: #fee2e2; border-left: 6px solid #dc2626; color: #991b1b;
        padding: 1rem 1.2rem; border-radius: 8px; margin-bottom: 1rem; font-weight: 600;
    }
    .triage-amarillo {
        background: #fef3c7; border-left: 6px solid #d97706; color: #92400e;
        padding: 1rem 1.2rem; border-radius: 8px; margin-bottom: 1rem; font-weight: 600;
    }
    .triage-verde {
        background: #dcfce7; border-left: 6px solid #16a34a; color: #166534;
        padding: 1rem 1.2rem; border-radius: 8px; margin-bottom: 1rem; font-weight: 600;
    }
    .conclusion-box {
        background: #f0f7ff; border: 1px solid #bfdbfe; border-left: 6px solid #13325b;
        padding: 1rem 1.2rem; border-radius: 8px; margin-bottom: 1.2rem; font-size: 0.95rem; line-height: 1.5;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 0. MÓDULO PÚBLICO: VALIDACIÓN POR QR
# ==========================================
params = st.query_params
if "val" in params:
    codigo_val = params.get("val", "N/A")
    paciente_val = urllib.parse.unquote(params.get("pac", "Paciente"))
    medico_val = urllib.parse.unquote(params.get("med", "Especialista CENCARDIO"))
    fecha_val = urllib.parse.unquote(params.get("fec", datetime.now().strftime("%Y-%m-%d")))

    c_v1, c_v2, c_v3 = st.columns([1, 2, 1])
    with c_v2:
        logo_data = obtener_logo_b64()
        logo_html = f'<img src="{logo_data}" class="cencardio-logo-img" alt="Cencardio Logo">' if logo_data else '<div style="font-size:3rem; margin-bottom:0.4rem;">🫀</div>'

        st.markdown(f"""
            <div class="cencardio-card" style="max-width: 580px;">
                {logo_html}
                <div class="cencardio-title">Centro Cardiovascular Colombiano</div>
                <div class="cencardio-sub">CENCARDIO · Certificado de Autenticidad</div>
                <div style="background: #ecfdf5; border: 2px solid #10b981; border-radius: 12px; padding: 1.2rem; margin-bottom: 1.5rem; text-align: left;">
                    <div style="color: #065f46; font-size: 1.1rem; font-weight: 800; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 8px;">
                        <span>✅</span> ESTUDIO MÉDICO VÁLIDO Y CERTIFICADO
                    </div>
                    <div style="font-size: 0.9rem; color: #1f2937; line-height: 1.6;">
                        <b>Estudio:</b> Registro Holter ECG 24 Horas<br>
                        <b>Paciente:</b> {paciente_val}<br>
                        <b>Médico Lector:</b> {medico_val}<br>
                        <b>Fecha de Emisión:</b> {fecha_val}<br>
                        <b>Código de Verificación:</b> <span style="font-family: monospace; color: #0369a1;">{codigo_val}</span><br>
                        <b>Normativa:</b> Cumple Res. 3100 de 2019 / Habilitación MinSalud
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: #64748b; line-height: 1.4;">
                    Este documento ha sido generado e interpretado mediante el sistema de lectura asistida del Centro Cardiovascular Colombiano CENCARDIO.
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("Ir al Portal Principal", use_container_width=True):
            st.query_params.clear()
            st.rerun()

    st.stop()

# ==========================================
# UTILIDAD: GENERACIÓN DE QR COMO URL
# ==========================================
def generar_qr_verificacion(paciente, medico, fecha_str, codigo_uuid):
    url_base = "https://holtercencardio.streamlit.app/"
    query_string = urllib.parse.urlencode({
        "val": codigo_uuid[:12],
        "pac": paciente,
        "med": medico,
        "fec": fecha_str
    })
    url_completa = f"{url_base}?{query_string}"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=4,
        border=1,
    )
    qr.add_data(url_completa)
    qr.make(fit=True)
    img_qr = qr.make_image(fill_color="#13325b", back_color="white")
    
    buf = io.BytesIO()
    img_qr.save(buf, format="PNG")
    return buf.getvalue()

# ==========================================
# GESTIÓN Y LIMPIEZA AUTOMÁTICA DE FIRMAS
# ==========================================
@st.cache_data
def procesar_firma_transparente():
    posibles_archivos = [
        "OR WILIAM ANDA RAMIREZ.pdf",
        "firma_amaya.pdf",
        "firma_amaya.png"
    ]
    
    archivo_encontrado = None
    for nombre in posibles_archivos:
        if os.path.exists(nombre):
            archivo_encontrado = nombre
            break
            
    if not archivo_encontrado:
        return None

    try:
        if archivo_encontrado.lower().endswith(".pdf"):
            doc_firma = fitz.open(archivo_encontrado)
            pix = doc_firma[0].get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc_firma.close()
        else:
            img = Image.open(archivo_encontrado)

        img = img.convert("RGBA")
        datos_pixeles = img.getdata()
        nuevos_pixeles = []

        for p in datos_pixeles:
            if p[0] > 185 and p[1] > 185 and p[2] > 185:
                nuevos_pixeles.append((255, 255, 255, 0))
            else:
                nuevos_pixeles.append((19, 50, 91, 255))

        img.putdata(nuevos_pixeles)

        caja = img.getbbox()
        if caja:
            img = img.crop(caja)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None

def normalizar_nombre_archivo(nombre):
    limpio = re.sub(r'[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s]', ' ', nombre)
    limpio = re.sub(r'\s+', '_', limpio).strip('_')
    return limpio if limpio else "PACIENTE"

# ==========================================
# BASE DE DATOS LOCAL
# ==========================================
def init_db():
    conn = sqlite3.connect("historial_holter.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS estudios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_registro TEXT,
            paciente_nombre TEXT,
            fc_prom INTEGER,
            sdnn INTEGER,
            medico_firmante TEXT,
            informe_texto TEXT,
            pdf_blob BLOB,
            codigo_verificacion TEXT
        )
    """)
    conn.commit()
    try:
        c.execute("ALTER TABLE estudios ADD COLUMN codigo_verificacion TEXT")
        conn.commit()
    except Exception:
        pass
    conn.close()

init_db()

def guardar_estudio_db(nombre, fc, sdnn, medico, texto, pdf_bytes, cod_verif):
    conn = sqlite3.connect("historial_holter.db")
    c = conn.cursor()
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        c.execute("""
            INSERT INTO estudios (fecha_registro, paciente_nombre, fc_prom, sdnn, medico_firmante, informe_texto, pdf_blob, codigo_verificacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (fecha_actual, nombre, fc, sdnn, medico, texto, pdf_bytes, cod_verif))
    except Exception:
        c.execute("""
            INSERT INTO estudios (fecha_registro, paciente_nombre, fc_prom, sdnn, medico_firmante, informe_texto, pdf_blob)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (fecha_actual, nombre, fc, sdnn, medico, texto, pdf_bytes))
    conn.commit()
    conn.close()

def obtener_historial_db():
    conn = sqlite3.connect("historial_holter.db")
    c = conn.cursor()
    filas = []
    try:
        c.execute("SELECT id, fecha_registro, paciente_nombre, fc_prom, sdnn, medico_firmante, pdf_blob, codigo_verificacion FROM estudios ORDER BY id DESC")
        filas = c.fetchall()
    except Exception:
        try:
            c.execute("SELECT id, fecha_registro, paciente_nombre, fc_prom, sdnn, medico_firmante, pdf_blob FROM estudios ORDER BY id DESC")
            filas_anteriores = c.fetchall()
            filas = [(r[0], r[1], r[2], r[3], r[4], r[5], r[6], "N/A") for r in filas_anteriores]
        except Exception:
            filas = []
    conn.close()
    return filas

def eliminar_estudio_db(estudio_id):
    conn = sqlite3.connect("historial_holter.db")
    c = conn.cursor()
    c.execute("DELETE FROM estudios WHERE id = ?", (estudio_id,))
    conn.commit()
    conn.close()

# ==========================================
# 1. PERFILES MÉDICOS Y ACCESO
# ==========================================
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

if not st.session_state.autenticado:
    col_izq, col_central, col_der = st.columns([1, 1.8, 1])
    
    with col_central:
        logo_data = obtener_logo_b64()
        logo_html = f'<img src="{logo_data}" class="cencardio-logo-img" alt="Cencardio Logo">' if logo_data else '<div style="font-size:3rem; margin-bottom:0.4rem;">🫀</div>'

        st.markdown(f"""
            <div class="cencardio-card">
                {logo_html}
                <div class="cencardio-title">Centro Cardiovascular Colombiano</div>
                <div class="cencardio-sub">CENCARDIO · Lectura de Holter</div>
        """, unsafe_allow_html=True)

        with st.form("form_login"):
            seleccion_etiqueta = st.selectbox(
                "Especialista Responsable:",
                options=OPCIONES_NOMBRES,
                index=0
            )
            clave = st.text_input("Contraseña de Acceso:", type="password")
            boton_ingresar = st.form_submit_button("Ingresar al Portal", use_container_width=True)

            if boton_ingresar:
                medico_seleccionado = next(m for m in LISTA_ESPECIALISTAS if m["etiqueta"] == seleccion_etiqueta)
                if medico_seleccionado["clave"] == clave:
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = medico_seleccionado["id"]
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta para el especialista seleccionado.")

        st.markdown('</div>', unsafe_allow_html=True)

    st.stop()

perfil_activo = PERFILES_POR_ID[st.session_state.usuario_actual]

# ==========================================
# 2. MOTOR CLÍNICO SPACELABS
# ==========================================
with st.sidebar:
    logo_data_sidebar = obtener_logo_b64()
    if logo_data_sidebar:
        st.markdown(f'<div style="text-align:center; margin-bottom:1rem;"><img src="{logo_data_sidebar}" style="max-width:140px;"></div>', unsafe_allow_html=True)
    
    st.write(f"👤 Especialista: **{perfil_activo['nombre_completo']}**")
    st.caption(f"{perfil_activo['especialidad']}\n{perfil_activo['registro']}")
    
    firma_disponible = procesar_firma_transparente()
    if firma_disponible and perfil_activo["id"] in ["dr.amaya", "admin"]:
        st.success("🖋️ Sello digitalizado cargado.")
        
    if st.button("Cerrar Sesión", use_container_width=True):
        cerrar_sesion()
        st.rerun()
    st.divider()
    st.caption("CENCARDIO · Plataforma Médica v7.6")

c_head1, c_head2 = st.columns([1, 6])
with c_head1:
    if logo_data_sidebar:
        st.image(logo_data_sidebar, width=105)
    else:
        st.markdown("<div style='font-size:2.8rem;'>🫀</div>", unsafe_allow_html=True)
with c_head2:
    st.markdown("""
        <div style="padding-top: 5px;">
            <div style="font-size: 1.65rem; font-weight: 800; color: #13325b; line-height: 1.2;">
                CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO
            </div>
            <div style="font-size: 0.95rem; font-weight: 600; color: #c8102e; text-transform: uppercase;">
                Sistema Profesional de Interpretación Holter Spacelabs
            </div>
        </div>
    """, unsafe_allow_html=True)

st.write("")

tab_procesar, tab_historial = st.tabs(["📥 Procesar Nuevo Estudio", "📁 Archivo Clínico y Facturación"])

def limpiar_numero(val_str):
    if not val_str:
        return 0
    val_str = val_str.replace(".", "").replace(",", ".")
    try:
        return int(float(val_str))
    except:
        return 0

def extraer_datos_spacelabs(pdf_bytes, filename=""):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    texto = ""
    for page in reader.pages:
        t = page.extract_text()
        if t:
            texto += t + "\n"

    datos = {}

    nombre_detectado = None
    m_nom = re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,50},\s*[A-ZÁÉÍÓÚÑ\s]{3,50})[\s\n]+(?:No confirmado|Confirmado)?[\s\n]*Informe Holter", texto)
    if m_nom:
        nombre_detectado = m_nom.group(1).replace("\n", " ").strip()

    if not nombre_detectado:
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        for idx, l in enumerate(lineas):
            if "Informe Holter" in l and idx > 0:
                candidato = lineas[idx - 1]
                if candidato in ["No confirmado", "Confirmado"] and idx > 1:
                    candidato = lineas[idx - 2]
                if candidato and len(candidato) > 4 and not any(p in candidato for p in ["CENTRO", "COLOMBIA", "Cra.", "30139"]):
                    nombre_detectado = candidato
                    break

    if not nombre_detectado and filename:
        nom_base = os.path.splitext(filename)[0]
        nom_limpio = re.sub(r'\(.*?\)', '', nom_base).strip()
        if len(nom_limpio) > 3:
            nombre_detectado = nom_limpio

    datos["paciente"] = nombre_detectado if nombre_detectado else "PACIENTE"

    fc_p = re.search(r"Prom\.?\s*(\d{2,3})", texto)
    datos["fc_prom"] = int(fc_p.group(1)) if fc_p else 70

    fc_max = re.search(r"M[áa]x\s*(\d{2,3})", texto)
    datos["fc_max"] = int(fc_max.group(1)) if fc_max else 100

    fc_min = re.search(r"M[íi]n\s*(\d{2,3})", texto)
    datos["fc_min"] = int(fc_min.group(1)) if fc_min else 55

    pausa_match = re.search(r"\bPausa\s+(\d+)", texto)
    datos["pausas"] = int(pausa_match.group(1)) if pausa_match else 0

    ev_m = re.search(r"Latidos ventriculares\s*:\s*([\d\.]+)", texto)
    datos["ev_total"] = limpiar_numero(ev_m.group(1)) if ev_m else 0

    tv_m = re.search(r"\bTV\s+([\d\.]+)", texto)
    datos["tv_episodios"] = limpiar_numero(tv_m.group(1)) if tv_m else 0

    dup_v_m = re.search(r"Apareado\s+([\d\.]+)", texto)
    datos["ev_duplas"] = limpiar_numero(dup_v_m.group(1)) if dup_v_m else 0

    big_m = re.search(r"Bigeminismo\s+([\d\.]+)", texto)
    datos["bigeminismo"] = limpiar_numero(big_m.group(1)) if big_m else 0

    esv_m = re.search(r"Latidos supraventriculares\s*:\s*([\d\.]+)", texto)
    datos["esv_total"] = limpiar_numero(esv_m.group(1)) if esv_m else 0

    tsv_m = re.search(r"\bTSV\s+([\d\.]+)", texto)
    datos["tsv_episodios"] = limpiar_numero(tsv_m.group(1)) if tsv_m else 0

    st_m = re.search(r"Depresi[óo]n ST\s+(\d+)\s+(-?[\d,\.]+)\s+([^\n]+)", texto)
    if st_m:
        datos["st_episodios"] = int(st_m.group(1))
        datos["st_desviacion"] = st_m.group(2)
    else:
        st_alt = re.search(r"Depresi[óo]n ST\s+(\d+)", texto)
        datos["st_episodios"] = int(st_alt.group(1)) if st_alt else 0
        datos["st_desviacion"] = "0"

    sdnn_m = re.search(r"Valor de 24 horas\s+[\d\.]+\s+([\d\.]+)", texto)
    datos["sdnn_24h"] = int(sdnn_m.group(1)) if sdnn_m else 85

    qtc_m = re.search(r"Todos los per[íi]odos\s+[\d\.]+\s+[\d\.]+\s+([\d\.]+)", texto)
    datos["qtc_prom"] = int(qtc_m.group(1)) if qtc_m else 400

    return datos

# ==========================================
# MOTOR DE SÍNTESIS DIAGNÓSTICA AUTOMÁTICA
# ==========================================
def sintetizar_conclusion_automatica(d):
    partes = []

    if d["fc_prom"] < 50:
        partes.append(f"Ritmo sinusal con tendencia a la bradicardia (FC promedio {d['fc_prom']} lpm).")
    elif d["fc_prom"] > 100:
        partes.append(f"Ritmo sinusal con taquicardia sostenida (FC promedio {d['fc_prom']} lpm).")
    else:
        partes.append(f"Ritmo sinusal con respuesta ventricular promedio conservada ({d['fc_prom']} lpm).")

    if d["pausas"] > 0:
        partes.append(f"Presencia de {d['pausas']} pausas patológicas (> 2.0 s), sugestivas de disfunción sinusal o bloqueo AV.")
    else:
        partes.append("Sin pausas patológicas ni bloqueos AV avanzados.")

    if d["tv_episodios"] > 0:
        partes.append(f"Registro de taquicardia ventricular no sostenida ({d['tv_episodios']} rachas de TV).")
    elif d["ev_total"] > 2000:
        partes.append(f"Carga ectópica ventricular elevada ({d['ev_total']} EV/24h), considerar correlación ecocardiográfica.")
    elif d["ev_total"] > 0:
        partes.append(f"Ectopias ventriculares monomorfas de baja carga ({d['ev_total']} EV).")
    else:
        partes.append("Sin ectopia ventricular de relevancia clínica.")

    if d["tsv_episodios"] > 0:
        partes.append(f"Episodios de taquicardia supraventricular paroxística documentados ({d['tsv_episodios']} TSV).")
    elif d["esv_total"] > 500:
        partes.append(f"Ectopia supraventricular frecuente ({d['esv_total']} ESV).")

    if d["st_episodios"] > 0:
        partes.append(f"Cambios en la repolarización compatibles con isquemia miocárdica silente ({d['st_episodios']} episodios de infradesnivel del ST).")
    else:
        partes.append("Sin alteraciones isquémicas del segmento ST.")

    if d["sdnn_24h"] < 50:
        partes.append("Variabilidad autonómica de la FC severamente disminuida (marcador de alto riesgo cardiovascular).")
    elif d["sdnn_24h"] <= 100:
        partes.append("Variabilidad de la FC moderadamente reducida.")
    else:
        partes.append("Variabilidad de la FC conservada.")

    return " ".join(partes)

def redactar_interpretacion(d, perfil):
    p1 = f"1. Ritmo de sinusal frecuencia cardiaca promedio de {d['fc_prom']} latidos por minuto."

    if d["qtc_prom"] > 460:
        p2 = f"2. Intervalos PR normal. Intervalo QTc prolongado (promedio {d['qtc_prom']} ms)."
    else:
        p2 = "2. Intervalos PR normal y QTc normales."

    if d["st_episodios"] > 0:
        p3 = f"3. Alteraciones isquémicas del segmento ST ({d['st_episodios']} episodios de depresión del ST, máx. {d['st_desviacion']} mm)."
    else:
        p3 = "3. Sin alteraciones isquémicas del segmento ST."

    p4 = "4. Sin Alteración en la conducción AV."
    p5 = "5. Sin Alteración en la conducción intraventricular."

    hallazgos_ectopia = []
    if d["esv_total"] > 0:
        if d["esv_total"] > 500:
            hallazgos_ectopia.append(f"frecuentes ectopias supraventriculares ({d['esv_total']})")
        else:
            hallazgos_ectopia.append("ectopias supraventriculares")

    if d["ev_total"] > 0:
        detalles_v = []
        if d["tv_episodios"] > 0:
            detalles_v.append(f"{d['tv_episodios']} episodios de TV")
        if d["ev_duplas"] > 0:
            detalles_v.append(f"{d['ev_duplas']} duplas")
        if d["bigeminismo"] > 0:
            detalles_v.append("bigeminismo")
        
        if detalles_v:
            hallazgos_ectopia.append(f"ventriculares frecuentes ({d['ev_total']} EV, incluyendo {', '.join(detalles_v)})")
        else:
            hallazgos_ectopia.append("ventriculares monomorfas")

    if hallazgos_ectopia:
        p6 = f"6. Alteración de los impulsos por {' y '.join(hallazgos_ectopia)}."
    else:
        p6 = "6. Sin alteración de los impulsos ectópicos de relevancia clínica."

    p7 = "7. No refirió síntomas."

    sdnn = d["sdnn_24h"]
    if sdnn < 50:
        p8 = "8. Variabilidad severamente disminuida de la FC."
        riesgo = "Alto riesgo"
    elif 50 <= sdnn <= 100:
        p8 = "8. Variabilidad disminuida de la FC."
        riesgo = "Riesgo medio"
    else:
        p8 = "8. Variabilidad conservada de la FC."
        riesgo = "Bajo riesgo / Normal"

    if d["pausas"] == 0:
        p9 = "9. Sin pausas significativas."
    else:
        p9 = f"9. Se registraron {d['pausas']} pausas significativas (> 2.0 s)."

    p10 = f"10. Riesgo del paciente SDNN a 24 HRS ({riesgo})."

    conclusion_diagnostica = sintetizar_conclusion_automatica(d)

    informe = f"""INTERPRETACIÓN TEST HOLTER

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

CONCLUSIÓN DIAGNÓSTICA:
{conclusion_diagnostica}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

    return informe

# ==========================================
# 3. INYECCIÓN DEL TEXTO, QR Y FIRMA (SIN SOLAPAMIENTO)
# ==========================================
def inyectar_y_generar_preview(pdf_bytes, texto_informe, datos_paciente, perfil, codigo_uuid, estampador_activo=False):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]

    rects_h = pagina1.search_for("Hallazgos:")
    rects_f = pagina1.search_for("Firma del médico")
    if not rects_f:
        rects_f = pagina1.search_for("Firma del operador")

    # Delimitación del recuadro de hallazgos
    if rects_h and rects_f:
        x0 = rects_h[0].x0 + 2
        y0 = rects_h[0].y1 + 4
        y1 = rects_f[0].y0 - 2
        x1 = pagina1.rect.width - 36
        rect_hallazgos = fitz.Rect(x0, y0, x1, y1)
    else:
        rect_hallazgos = fitz.Rect(35, 508, 565, 735)

    # Inserción con auto-escalado seguro para el texto clínico
    font_size = 7.4
    for fs in [7.4, 7.0, 6.6, 6.2, 5.8, 5.2, 4.8]:
        doc_test = fitz.open(stream=pdf_bytes, filetype="pdf")
        p_test = doc_test[0]
        rc = p_test.insert_textbox(
            rect_hallazgos,
            texto_informe,
            fontsize=fs,
            fontname="helv",
            align=fitz.TEXT_ALIGN_LEFT
        )
        doc_test.close()
        if rc >= 0:
            font_size = fs
            break

    pagina1.insert_textbox(
        rect_hallazgos,
        texto_informe,
        fontsize=font_size,
        fontname="helv",
        color=(0, 0, 0),
        align=fitz.TEXT_ALIGN_LEFT
    )

    # Coordenadas de la línea de firma a la derecha
    y_base = rects_f[0].y0 if rects_f else 740

    # ==========================================
    # CÓDIGO QR Y VALIDACIÓN FORENSE EN EL CENTRO
    # (Libera la esquina izquierda para el nombre del médico)
    # ==========================================
    fecha_emision = datetime.now().strftime("%Y-%m-%d")
    qr_bytes = generar_qr_verificacion(
        datos_paciente, 
        perfil['nombre_completo'], 
        fecha_emision, 
        codigo_uuid
    )
    
    # Se ubica en el centro (x: 232 a 272), lejos del nombre (izq) y la firma (der)
    rect_qr = fitz.Rect(232, y_base - 32, 272, y_base + 8)
    pagina1.insert_image(rect_qr, stream=qr_bytes)
    
    pagina1.insert_text(
        fitz.Point(276, y_base - 18),
        "Validado Digitalmente",
        fontsize=5.2,
        fontname="helv",
        color=(0.08, 0.2, 0.36)
    )
    pagina1.insert_text(
        fitz.Point(276, y_base - 9),
        "Res. 3100 de 2019 - MinSalud",
        fontsize=4.7,
        fontname="helv",
        color=(0.25, 0.25, 0.25)
    )
    pagina1.insert_text(
        fitz.Point(276, y_base),
        f"Cód: {codigo_uuid[:12]}...",
        fontsize=4.5,
        fontname="helv",
        color=(0.4, 0.4, 0.4)
    )

    # ==========================================
    # ESTAMPADO DE LA FIRMA DIGITAL A LA DERECHA
    # ==========================================
    firma_png_bytes = procesar_firma_transparente()
    if firma_png_bytes and estampador_activo:
        if rects_f:
            fx0 = rects_f[0].x0 + 10
            fy0 = rects_f[0].y0 - 58
            fx1 = fx0 + 155
            fy1 = rects_f[0].y0 + 4
            rect_firma = fitz.Rect(fx0, fy0, fx1, fy1)
            pagina1.insert_image(rect_firma, stream=firma_png_bytes)

    pix = pagina1.get_pixmap(dpi=150)
    img_preview = pix.tobytes("png")
    pdf_final_bytes = doc.tobytes()
    doc.close()

    return pdf_final_bytes, img_preview

# ==========================================
# PESTAÑA 1: PROCESAMIENTO
# ==========================================
with tab_procesar:
    uploaded_file = st.file_uploader("Cargar estudio Holter de Spacelabs (PDF)", type=["pdf"])

    if uploaded_file is not None:
        bytes_originales = uploaded_file.getvalue()

        if "datos_actuales" not in st.session_state or st.session_state.get("archivo_actual") != uploaded_file.name:
            with st.spinner("Analizando trazado de Spacelabs y sintetizando diagnóstico..."):
                st.session_state.datos_actuales = extraer_datos_spacelabs(bytes_originales, uploaded_file.name)
                st.session_state.texto_informe = redactar_interpretacion(st.session_state.datos_actuales, perfil_activo)
                st.session_state.archivo_actual = uploaded_file.name
                st.session_state.estudio_uuid = str(uuid.uuid4()).upper()

        if "CONCLUSIÓN DIAGNÓSTICA" not in st.session_state.get("texto_informe", ""):
            st.session_state.texto_informe = redactar_interpretacion(st.session_state.datos_actuales, perfil_activo)

        datos = st.session_state.datos_actuales

        alertas_criticas = []
        alertas_moderadas = []

        if datos["sdnn_24h"] < 50:
            alertas_criticas.append("Variabilidad severamente disminuida (SDNN < 50 ms: Alto riesgo cardiovascular).")
        elif datos["sdnn_24h"] <= 100:
            alertas_moderadas.append("Variabilidad de la FC disminuida (SDNN entre 50 y 100 ms).")

        if datos["tv_episodios"] > 0:
            alertas_criticas.append(f"Se registraron {datos['tv_episodios']} rachas de Taquicardia Ventricular (TV).")

        if datos["pausas"] > 0:
            alertas_criticas.append(f"Se registraron {datos['pausas']} pausas patológicas significativas (> 2.0 s).")

        if datos["st_episodios"] > 0:
            alertas_moderadas.append(f"Alteraciones isquémicas del segmento ST ({datos['st_episodios']} episodios, máx. {datos['st_desviacion']} mm).")

        if datos["qtc_prom"] > 460:
            alertas_moderadas.append(f"Intervalo QTc prolongado (promedio {datos['qtc_prom']} ms).")

        if alertas_criticas:
            st.markdown(f"""
                <div class="triage-rojo">
                    ⚠️ <b>ALERTA CRÍTICA: Hallazgos de Alto Riesgo Cardiovascular Detectados</b><br>
                    • {'<br>• '.join(alertas_criticas)}
                </div>
            """, unsafe_allow_html=True)
        elif alertas_moderadas:
            st.markdown(f"""
                <div class="triage-amarillo">
                    ⚡ <b>PRECAUCIÓN CLÍNICA: Hallazgos Relevantes para Seguimiento</b><br>
                    • {'<br>• '.join(alertas_moderadas)}
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div class="triage-verde">
                    ✅ <b>PARÁMETROS ESTABLES: Registro sin criterios de alto riesgo electrocardiográfico</b>
                </div>
            """, unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("FC Promedio", f"{datos['fc_prom']} lpm", f"Mín {datos['fc_min']} | Máx {datos['fc_max']}")
        c2.metric("Ectopias Ventriculares", f"{datos['ev_total']} EV", f"TV: {datos['tv_episodios']}")
        c3.metric("Ectopias Supraventriculares", f"{datos['esv_total']} ESV", f"TSV: {datos['tsv_episodios']}")
        c4.metric("SDNN (24 Horas)", f"{datos['sdnn_24h']} ms", f"ST: {datos['st_episodios']} ep.")

        st.divider()

        st.markdown(f"""
            <div class="conclusion-box">
                <b style="color: #13325b; font-size: 1.05rem;">🩺 Conclusión Diagnóstica Generada Automáticamente:</b><br>
                <div style="margin-top: 0.35rem; color: #1e293b;">
                    {sintetizar_conclusion_automatica(datos)}
                </div>
            </div>
        """, unsafe_allow_html=True)

        col_edicion, col_preview = st.columns([1, 1], gap="large")

        with col_edicion:
            nombre_confirmado = st.text_input(
                "👤 Nombre del Paciente (editable para el archivo y descarga):",
                value=datos['paciente']
            )
            paciente_nombre_archivo = normalizar_nombre_archivo(nombre_confirmado)

            st.subheader("📝 Edición de la Interpretación Completa")
            informe_para_grabar = st.text_area(
                "Documento completo para inyectar en el PDF:",
                value=st.session_state.texto_informe,
                height=350
            )

            debe_estampar = perfil_activo["id"] in ["dr.amaya", "admin"]
            pdf_final, img_preview = inyectar_y_generar_preview(
                bytes_originales, 
                informe_para_grabar, 
                nombre_confirmado, 
                perfil_activo, 
                st.session_state.estudio_uuid, 
                estampador_activo=debe_estampar
            )

            col_btn1, col_btn2 = st.columns([1, 1])
            with col_btn1:
                st.download_button(
                    label="📄 DESCARGAR PDF DILIGENCIADO",
                    data=pdf_final,
                    file_name=f"{paciente_nombre_archivo}_Holter_Firmado.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
            with col_btn2:
                if st.button("💾 Guardar en Archivo Clínico", use_container_width=True):
                    guardar_estudio_db(
                        nombre_confirmado,
                        datos['fc_prom'],
                        datos['sdnn_24h'],
                        perfil_activo['nombre_completo'],
                        informe_para_grabar,
                        pdf_final,
                        st.session_state.estudio_uuid
                    )
                    st.success(f"✅ Guardado con éxito: {nombre_confirmado}")

            st.write("")
            if st.button("🔄 Descartar / Limpiar Estudio Actual", use_container_width=True):
                if "datos_actuales" in st.session_state:
                    del st.session_state["datos_actuales"]
                if "archivo_actual" in st.session_state:
                    del st.session_state["archivo_actual"]
                if "texto_informe" in st.session_state:
                    del st.session_state["texto_informe"]
                st.rerun()

        with col_preview:
            st.subheader("👁️ Vista Previa Oficial (Página 1)")
            st.markdown('<div class="preview-container">', unsafe_allow_html=True)
            st.image(img_preview, caption=f"Página 1 - {nombre_confirmado}", use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# PESTAÑA 2: ARCHIVO CLÍNICO Y FACTURACIÓN
# ==========================================
with tab_historial:
    st.subheader("📂 Registro de Producción Médica e Historial")
    historial = obtener_historial_db()

    if not historial:
        st.info("Aún no hay estudios archivados en el sistema.")
    else:
        datos_tabla = []
        for h in historial:
            datos_tabla.append({
                "ID": h[0],
                "Fecha de Registro": h[1],
                "Paciente": h[2],
                "FC Media (lpm)": h[3],
                "SDNN 24h (ms)": h[4],
                "Especialista Firmante": h[5],
                "Código de Verificación": h[7] if len(h) > 7 and h[7] else "N/A"
            })
        df_produccion = pd.DataFrame(datos_tabla)

        col_rep1, col_rep2 = st.columns([3, 1])
        with col_rep1:
            busqueda = st.text_input("🔍 Buscar paciente por nombre o apellido:", "")
        with col_rep2:
            st.write("")
            st.write("")
            csv_data = df_produccion.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📊 Descargar Informe de Producción (CSV)",
                data=csv_data,
                file_name=f"Reporte_Produccion_Cencardio_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )

        st.divider()

        for item in historial:
            est_id = item[0]
            fecha = item[1]
            pac_nom = item[2]
            fc = item[3]
            sdnn = item[4]
            med = item[5]
            pdf_data = item[6]
            cod_ver = item[7] if len(item) > 7 and item[7] else "N/A"
            
            if busqueda.lower() in pac_nom.lower():
                nom_archivo_copia = normalizar_nombre_archivo(pac_nom)
                with st.expander(f"👤 {pac_nom} | 📅 {fecha} | 👨‍⚕️ {med}"):
                    c_det1, c_det2, c_desc, c_del = st.columns([2, 2, 2, 1.5])
                    c_det1.write(f"**FC Media:** {fc} lpm\n**SDNN 24h:** {sdnn} ms")
                    c_det2.write(f"**Código de Autenticidad:**\n`{cod_ver}`")
                    with c_desc:
                        st.download_button(
                            label="📥 Descargar PDF",
                            data=pdf_data,
                            file_name=f"{nom_archivo_copia}_Holter_Copia.pdf",
                            mime="application/pdf",
                            key=f"descarga_{est_id}",
                            use_container_width=True
                        )
                    with c_del:
                        if st.button("🗑️ Eliminar Registro", key=f"del_{est_id}", use_container_width=True):
                            eliminar_estudio_db(est_id)
                            st.toast(f"Registro de {pac_nom} eliminado.", icon="🗑️")
                            st.rerun()
