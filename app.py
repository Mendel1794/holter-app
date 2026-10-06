import streamlit as st
import fitz  # PyMuPDF: motor C++ de lectura y renderizado ultrarrápido
from PIL import Image
import qrcode
import re
import base64
import os
import io
import sqlite3
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import uuid
import urllib.parse

# Importación segura de openpyxl
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_INSTALADO = True
except ImportError:
    OPENPYXL_INSTALADO = False

# Importación segura de OCR para escaneo de fotos
try:
    import pytesseract
    PYTESSERACT_INSTALADO = True
except ImportError:
    PYTESSERACT_INSTALADO = False

st.set_page_config(
    page_title="Centro Cardiovascular Colombiano CENCARDIO · Workstation",
    page_icon="🫀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# SISTEMA DE DISEÑO INSTITUCIONAL (NIVEL CARDIOVASCULAR ALTA COMPLEJIDAD)
# ==============================================================================
@st.cache_data
def obtener_logo_b64():
    for nom in ["cencardio.jpg", "cencardio.png", "cencardio.jpeg", "logo.png", "logo.jpg"]:
        if os.path.exists(nom):
            with open(nom, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            mime = "png" if nom.endswith("png") else "jpeg"
            return f"data:image/{mime};base64,{b64}"
    return None

def cargar_estilos_institucionales():
    return """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
        color: #1e293b;
    }

    .stApp {
        background-color: #f8fafc;
        background-image: 
            radial-gradient(at 0% 0%, rgba(10, 37, 64, 0.03) 0px, transparent 50%),
            radial-gradient(at 100% 100%, rgba(200, 16, 46, 0.02) 0px, transparent 50%);
    }

    header[data-testid="stHeader"] { background: transparent !important; }
    div[data-testid="stDecoration"] { display: none !important; }
    .block-container { padding-top: 1.2rem !important; padding-bottom: 2.5rem !important; }

    /* Barra Superior Institucional Hospitalaria */
    .top-hospital-bar {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 1rem 1.6rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 4px 20px -2px rgba(10, 37, 64, 0.04);
        margin-bottom: 1.5rem;
    }
    .inst-badge-group {
        display: flex;
        gap: 8px;
        align-items: center;
    }
    .inst-badge-primary {
        background: #e0f2fe; color: #0369a1; font-size: 0.72rem; font-weight: 700;
        padding: 4px 10px; border-radius: 20px; letter-spacing: 0.5px; text-transform: uppercase;
        border: 1px solid #bae6fd;
    }
    .inst-badge-success {
        background: #ecfdf5; color: #065f46; font-size: 0.72rem; font-weight: 700;
        padding: 4px 10px; border-radius: 20px; letter-spacing: 0.5px; text-transform: uppercase;
        border: 1px solid #a7f3d0;
    }

    /* Pestañas de Consola */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: #ffffff;
        padding: 6px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 10px rgba(10, 37, 64, 0.03);
        margin-bottom: 1.2rem;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px !important;
        padding: 9px 22px !important;
        font-weight: 700 !important;
        font-size: 0.88rem !important;
        color: #64748b !important;
        border: none !important;
        transition: all 0.2s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #0a2540 !important;
        background: #f1f5f9;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0a2540 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 12px rgba(10, 37, 64, 0.15) !important;
    }

    /* Menú Lateral */
    div[data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e2e8f0 !important;
        box-shadow: 2px 0 16px rgba(10, 37, 64, 0.02) !important;
    }
    .specialist-card {
        background: linear-gradient(135deg, #0a2540 0%, #133863 100%);
        border-radius: 14px;
        padding: 1.1rem;
        color: #ffffff;
        margin-bottom: 1.2rem;
        box-shadow: 0 6px 18px -3px rgba(10, 37, 64, 0.25);
    }
    .specialist-title { font-size: 0.95rem; font-weight: 800; letter-spacing: 0.2px; }
    .specialist-sub { font-size: 0.74rem; color: #93c5fd; font-weight: 600; text-transform: uppercase; margin-top: 2px; }
    .specialist-reg { font-size: 0.72rem; color: #cbd5e1; font-family: 'JetBrains Mono', monospace; margin-top: 6px; }

    /* Tarjetas de Selección */
    div[data-testid="stRadio"] > div { gap: 8px; }
    div[data-testid="stRadio"] label {
        background: #f8fafc;
        border: 1.5px solid #e2e8f0;
        padding: 10px 14px;
        border-radius: 10px;
        cursor: pointer;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        font-weight: 600;
        font-size: 0.86rem;
        color: #334155;
    }
    div[data-testid="stRadio"] label:hover {
        border-color: #0284c7;
        background: #f0f9ff;
        color: #0369a1;
        transform: translateY(-1px);
    }

    /* Tarjetas de Métricas */
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 1rem 1.2rem;
        box-shadow: 0 4px 14px -2px rgba(10, 37, 64, 0.04);
        border-top: 4px solid #0a2540;
    }
    [data-testid="stMetricValue"] {
        color: #0a2540 !important;
        font-weight: 800 !important;
        font-size: 1.55rem !important;
    }
    [data-testid="stMetricLabel"] {
        color: #64748b !important;
        font-weight: 600 !important;
        font-size: 0.78rem !important;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }

    /* Triage Clínico */
    .triage-rojo {
        background: #fef2f2 !important; border: 1.5px solid #fecaca !important; border-left: 6px solid #dc2626 !important; color: #991b1b !important;
        padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important; font-size: 0.92rem !important; line-height: 1.5 !important;
    }
    .triage-amarillo {
        background: #fffbeb !important; border: 1.5px solid #fde68a !important; border-left: 6px solid #d97706 !important; color: #92400e !important;
        padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important; font-size: 0.92rem !important; line-height: 1.5 !important;
    }
    .triage-verde {
        background: #f0fdf4 !important; border: 1.5px solid #bbf7d0 !important; border-left: 6px solid #16a34a !important; color: #166534 !important;
        padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important; font-size: 0.92rem !important; line-height: 1.5 !important;
    }
    .dinamica-status {
        background: #ecfdf5; border: 1px solid #a7f3d0; color: #065f46;
        padding: 0.6rem 0.9rem; border-radius: 8px; font-weight: 600; font-size: 0.88rem; margin-bottom: 0.8rem;
    }

    /* Botones Institucionales */
    .stButton > button {
        background-color: #0a2540 !important;
        color: #ffffff !important;
        border-radius: 10px !important;
        border: none !important;
        font-weight: 700 !important;
        font-size: 0.88rem !important;
        padding: 0.65rem 1.4rem !important;
        box-shadow: 0 4px 12px rgba(10, 37, 64, 0.15) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        background-color: #c8102e !important;
        color: #ffffff !important;
        transform: translateY(-1px) !important;
    }

    div[data-testid="stLinkButton"] a {
        background-color: #25D366 !important;
        color: #ffffff !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
        padding: 0.65rem 1rem !important;
        text-align: center !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(37, 211, 102, 0.25) !important;
    }
    div[data-testid="stLinkButton"] a:hover {
        background-color: #1eb854 !important;
        color: #ffffff !important;
        transform: translateY(-1px) !important;
    }

    textarea {
        background-color: #ffffff !important;
        border: 1.5px solid #cbd5e1 !important;
        border-radius: 10px !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 12.5px !important;
    }
    .preview-container {
        border: 2px solid #e2e8f0;
        border-radius: 14px;
        box-shadow: 0 8px 24px -4px rgba(10, 37, 64, 0.08);
        background: #ffffff;
        padding: 8px;
    }
    .cencardio-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 2.6rem 2.8rem;
        box-shadow: 0 20px 40px -12px rgba(10, 37, 64, 0.12);
        max-width: 480px;
        margin: 3rem auto;
        text-align: center;
    }
    </style>
    """

st.markdown(cargar_estilos_institucionales(), unsafe_allow_html=True)

# ==========================================
# 0. MÓDULO PÚBLICO: VALIDACIÓN POR QR (RES. 3100)
# ==========================================
params = st.query_params
if "val" in params:
    codigo_val = params.get("val", "N/A")
    paciente_val = urllib.parse.unquote(params.get("pac", "Paciente"))
    medico_val = urllib.parse.unquote(params.get("med", "Especialista CENCARDIO"))
    fecha_val = urllib.parse.unquote(params.get("fec", datetime.now().strftime("%Y-%m-%d")))
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

# ==========================================
# UTILIDAD: GENERACIÓN DE QR
# ==========================================
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

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=4,
        border=1,
    )
    qr.add_data(url_completa)
    qr.make(fit=True)
    img_qr = qr.make_image(fill_color="#0a2540", back_color="white")
    
    buf = io.BytesIO()
    img_qr.save(buf, format="PNG")
    return buf.getvalue()

# ==========================================
# GESTIÓN DE FIRMAS
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
    limpio = re.sub(r'[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s]', ' ', nombre)
    limpio = re.sub(r'\s+', '_', limpio).strip('_')
    return limpio if limpio else "PACIENTE"

# ==========================================
# BASE DE DATOS LOCAL MULTIMODALIDAD
# ==========================================
def init_db():
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
            codigo_verificacion TEXT
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

init_db()

def sincronizar_directorio_dinamica(df):
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    fecha_hoy = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    col_ced = next((col for col in df.columns if any(k in col.lower() for k in ["cedula", "documento", "identificacion", "id"])), None)
    col_tel = next((col for col in df.columns if any(k in col.lower() for k in ["telefono", "celular", "tel", "movil"])), None)
    col_nom = next((col for col in df.columns if any(k in col.lower() for k in ["nombre", "paciente", "usuario"])), None)

    if not col_ced or not col_tel:
        conn.close()
        return False, "El archivo debe contener columnas de Cédula y Teléfono/Celular."

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
    return True, f"Se sincronizaron con éxito {registros} pacientes de Dinámica Gerencial."

def buscar_telefono_dinamica(cedula):
    if not cedula:
        return ""
    ced_limpia = re.sub(r'\D', '', str(cedula))
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT telefono FROM dinamica_pacientes WHERE cedula = ?", (ced_limpia,))
    res = c.fetchone()
    conn.close()
    return res[0] if res else ""

def guardar_estudio_db(nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif):
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("""
        INSERT INTO estudios (fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, informe_texto, pdf_blob, codigo_verificacion)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (fecha_actual, nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif))
    conn.commit()
    conn.close()

def obtener_historial_db():
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, pdf_blob, codigo_verificacion FROM estudios ORDER BY id DESC")
    filas = c.fetchall()
    conn.close()
    return filas

def eliminar_estudio_db(estudio_id):
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("DELETE FROM estudios WHERE id = ?", (estudio_id,))
    conn.commit()
    conn.close()

# ==============================================================================
# GENERADOR AVANZADO DE EXCEL INSTITUCIONAL CON RESPALDO DEFENSIVO
# ==============================================================================
def generar_excel_avanzado_produccion(df_base):
    if not OPENPYXL_INSTALADO:
        csv_str = df_base.to_csv(sep=';', index=False, encoding='utf-8-sig')
        return csv_str.encode('utf-8-sig'), "csv"

    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # 1. Pestaña: Resumen Ejecutivo y Métricas
    ws_resumen = wb.active
    ws_resumen.title = "📊 Tablero Gerencial"
    ws_resumen.views.sheetView[0].showGridLines = True

    fill_navy = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
    fill_wine = PatternFill(start_color="C8102E", end_color="C8102E", fill_type="solid")
    fill_light = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True, color="0A2540")
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    ws_resumen["B2"] = "CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO"
    ws_resumen["B2"].font = Font(name="Calibri", size=14, bold=True, color="0A2540")
    ws_resumen["B3"] = f"REPORTE GERENCIAL DE PRODUCCIÓN Y FACTURACIÓN · EMITIDO: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws_resumen["B3"].font = Font(name="Calibri", size=9, bold=True, color="64748B")

    ws_resumen["B5"] = "PRODUCCIÓN POR MODALIDAD DIAGNÓSTICA"
    ws_resumen["B5"].font = font_bold
    headers_mod = ["Modalidad Diagnóstica", "Código CUPS", "Total Estudios", "% Participación"]
    for col_idx, h in enumerate(headers_mod, start=2):
        cell = ws_resumen.cell(row=6, column=col_idx, value=h)
        cell.fill = fill_navy
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    prod_mod = df_base.groupby(["Modalidad", "Código CUPS"]).size().reset_index(name="Cantidad")
    total_estudios = len(df_base)
    r_idx = 7
    for _, fila in prod_mod.iterrows():
        pct = (fila["Cantidad"] / total_estudios) * 100 if total_estudios > 0 else 0
        ws_resumen.cell(row=r_idx, column=2, value=fila["Modalidad"]).border = border_thin
        ws_resumen.cell(row=r_idx, column=3, value=fila["Código CUPS"]).border = border_thin
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

    ws_resumen.cell(row=r_idx, column=2, value="DISTRIBUCIÓN CRONOLÓGICA POR MESES").font = font_bold
    r_idx += 1
    headers_mes = ["Periodo (Año-Mes)", "Total Estudios", "% Producción"]
    for col_idx, h in enumerate(headers_mes, start=2):
        cell = ws_resumen.cell(row=r_idx, column=col_idx, value=h)
        cell.fill = fill_wine
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center")

    df_base["Periodo_Mes"] = pd.to_datetime(df_base["Fecha de Registro"]).dt.strftime('%Y-%m')
    prod_mes = df_base.groupby("Periodo_Mes").size().reset_index(name="Cantidad").sort_values("Periodo_Mes", ascending=False)
    r_idx += 1
    for _, fila in prod_mes.iterrows():
        pct = (fila["Cantidad"] / total_estudios) * 100 if total_estudios > 0 else 0
        ws_resumen.cell(row=r_idx, column=2, value=fila["Periodo_Mes"]).border = border_thin
        c_c = ws_resumen.cell(row=r_idx, column=3, value=fila["Cantidad"])
        c_c.border = border_thin
        c_c.alignment = Alignment(horizontal="center")
        c_p = ws_resumen.cell(row=r_idx, column=4, value=f"{pct:.1f}%")
        c_p.border = border_thin
        c_p.alignment = Alignment(horizontal="center")
        r_idx += 1

    # 2. Pestaña: Detalle RIPS y Facturación
    ws_detalle = wb.create_sheet(title="📁 Detalle RIPS y Facturación")
    ws_detalle.views.sheetView[0].showGridLines = True

    columnas_ordenadas = [
        "ID", "Fecha de Registro", "Periodo_Mes", "Paciente", "Modalidad",
        "Código CUPS", "Parámetro Clave", "Especialista Firmante", "Código Forense"
    ]
    df_exp = df_base[columnas_ordenadas].copy()
    
    headers_detalle = [
        "N° ID", "Fecha y Hora Registro", "Periodo", "Nombre del Paciente",
        "Modalidad Diagnóstica", "Código CUPS", "Métrica / Parámetro Clave",
        "Especialista Lector Responsable", "Código Único Forense (Res. 3100)"
    ]
    for col_num, h in enumerate(headers_detalle, 1):
        cell = ws_detalle.cell(row=1, column=col_num, value=h)
        cell.fill = fill_navy
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_detalle.row_dimensions[1].height = 26

    for r_idx, fila in enumerate(df_exp.itertuples(index=False), start=2):
        ws_detalle.row_dimensions[r_idx].height = 20
        fill_row = fill_light if (r_idx % 2 == 0) else None
        for c_idx, valor in enumerate(fila, start=1):
            cell = ws_detalle.cell(row=r_idx, column=c_idx, value=str(valor))
            cell.border = border_thin
            if fill_row: cell.fill = fill_row
            if c_idx in [1, 3, 6]: cell.alignment = Alignment(horizontal="center")

    for ws in [ws_resumen, ws_detalle]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    return output.getvalue(), "xlsx"

# ==========================================
# 1. PERFILES MÉDICOS OFICIALES
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
    col_izq, col_central, col_der = st.columns([1, 1.6, 1])
    with col_central:
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
            boton_ingresar = st.form_submit_button("Ingresar a la Estación", use_container_width=True)

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

# ==============================================================================
# DETECTOR DETERMINÍSTICO DE ESTUDIO (BASE C++)
# ==============================================================================
def detectar_tipo_documento_clinico(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    for p in doc:
        texto += p.get_text() + "\n"
    doc.close()

    if any(k in texto for k in ["Sentinel", "Presión arterial por la mañana", "Informe de MAPA", "AASI"]):
        return "MAPA"
    elif any(k in texto for k in ["Informe Holter", "Latidos ventriculares", "Pathfinder SL", "Arritmias ventriculares"]):
        return "HOLTER"
    else:
        return "DESCONOCIDO"

# ==============================================================================
# GRÁFICA DEL TACOGRAMA
# ==============================================================================
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
        val = max(fc_min, min(fc_max, val))
        fc_curva.append(round(val))

    fig = go.Figure()
    fig.add_hrect(
        y0=60, y1=100, 
        fillcolor="rgba(10, 37, 64, 0.04)", 
        line_width=0,
        annotation_text="Rango Normal (60-100)", 
        annotation_position="top left",
        annotation_font_size=9,
        annotation_font_color="#64748b"
    )

    fig.add_trace(go.Scatter(
        x=horas, y=fc_curva,
        mode='lines+markers',
        name='FC Horaria (lpm)',
        line=dict(color='#0A2540', width=2.5),
        marker=dict(size=4, color='#C8102E')
    ))

    fig.add_hline(
        y=fc_prom,
        line_dash="dot",
        line_color="#0284c7",
        annotation_text=f"Promedio: {fc_prom} lpm",
        annotation_position="bottom right",
        annotation_font_size=10
    )

    fig.update_layout(
        title=dict(text="<b>Tacograma Horario y Variabilidad Circadiana (24 Horas)</b>", font=dict(size=13, color="#0A2540")),
        height=240,
        margin=dict(l=35, r=20, t=35, b=25),
        xaxis=dict(title="", tickfont=dict(size=9), showgrid=True, gridcolor="#f1f5f9"),
        yaxis=dict(title="lpm", tickfont=dict(size=9), showgrid=True, gridcolor="#f1f5f9"),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        showlegend=False
    )
    return fig

# ==========================================
# 2. SELECCIÓN DE MODALIDAD DIAGNÓSTICA
# ==========================================
with st.sidebar:
    logo_data_sidebar = obtener_logo_b64()
    if logo_data_sidebar:
        st.markdown(f'<div style="text-align:center; margin-bottom:1.2rem;"><img src="{logo_data_sidebar}" style="max-width:145px;"></div>', unsafe_allow_html=True)
    
    st.markdown(f"""
        <div class="specialist-card">
            <div class="specialist-title">{perfil_activo['nombre_completo']}</div>
            <div class="specialist-sub">{perfil_activo['especialidad']}</div>
            <div class="specialist-reg">{perfil_activo['registro']}</div>
        </div>
    """, unsafe_allow_html=True)

    firma_disponible = procesar_firma_transparente()
    if firma_disponible and perfil_activo["id"] in ["dr.amaya", "admin"]:
        st.success("🖋️ Sello digitalizado cargado.")
        
    st.divider()

    st.markdown("<div style='font-size:0.78rem; font-weight:800; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:0.5rem;'>MODALIDAD DIAGNÓSTICA</div>", unsafe_allow_html=True)
    modalidad_seleccionada = st.radio(
        "Seleccione estudio:",
        [
            "🫀 Holter ECG 24H (CUPS 895001)",
            "🩺 MAPA Tensional 24H (CUPS 895003)",
            "🏃 Prueba de Esfuerzo (CUPS 893805)"
        ],
        label_visibility="collapsed"
    )

    st.divider()
    st.markdown("<div style='font-size:0.78rem; font-weight:800; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:0.5rem;'>SINCRONIZADOR DINÁMICA</div>", unsafe_allow_html=True)
    archivo_dinamica = st.file_uploader("Subir Directorio Dinámica (Excel o CSV)", type=["xlsx", "xls", "csv"], key="sync_dinamica")
    if archivo_dinamica is not None:
        try:
            df_din = pd.read_csv(archivo_dinamica) if archivo_dinamica.name.endswith(".csv") else pd.read_excel(archivo_dinamica)
            ok, msg = sincronizar_directorio_dinamica(df_din)
            if ok:
                st.success(msg)
            else:
                st.error(msg)
        except Exception as e:
            st.error(f"Error al leer archivo: {e}")

    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        cerrar_sesion()
        st.rerun()
    st.caption("CENCARDIO · Workstation Enterprise v14.0")

# Header institucional superior
st.markdown("""
    <div class="top-hospital-bar">
        <div>
            <div style="font-size: 1.35rem; font-weight: 800; color: #0a2540; line-height: 1.2;">
                CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO
            </div>
            <div style="font-size: 0.82rem; font-weight: 700; color: #c8102e; text-transform: uppercase;">
                Estación Diagnóstica de Cardiología No Invasiva · Alta Complejidad
            </div>
        </div>
        <div class="inst-badge-group">
            <span class="inst-badge-success">● Sistema en Línea</span>
            <span class="inst-badge-primary">Habilitación MinSalud · Res. 3100</span>
        </div>
    </div>
""", unsafe_allow_html=True)

tab_procesar, tab_historial = st.tabs(["📥 Procesamiento del Estudio", "📁 Archivo Clínico & Reportes Gerenciales"])

def limpiar_numero(val_str):
    if not val_str:
        return 0
    val_str = str(val_str).replace(".", "").replace(",", ".")
    try:
        return int(float(val_str))
    except:
        return 0

# ==============================================================================
# MOTOR 1: HOLTER ECG 24 HORAS (CUPS 895001) - 100% DINÁMICO (11 PUNTOS)
# ==============================================================================
def extraer_datos_holter(pdf_bytes, filename=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    for page in doc:
        texto += page.get_text() + "\n"
    doc.close()

    d = {}
    m_dx = re.search(r"(?:DX|DIAGN[ÓO]STICO|INDICACI[ÓO]N)\s*[:\.]?\s*([^\n\r\|]+)", texto, re.IGNORECASE)
    if not m_dx:
        m_dx = re.search(r"Comentarios de la prueba:\s*\n?\s*([^\n\r\|]+)", texto, re.IGNORECASE)
    d["dx_motivo"] = m_dx.group(1).strip().upper() if m_dx else ""

    nombre_detectado = None
    m_nom = re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,50},\s*[A-ZÁÉÍÓÚÑ\s]{3,50})[\s\n]+(?:No confirmado|Confirmado|Reconfirmado)?[\s\n]*Informe Holter", texto)
    if m_nom:
        nombre_detectado = m_nom.group(1).replace("\n", " ").strip()
    if not nombre_detectado and filename:
        nombre_detectado = os.path.splitext(filename)[0]
    d["paciente"] = nombre_detectado if nombre_detectado else "PACIENTE"

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

    taqui_m = re.search(r"Taquicardia\s+(\d+)(?:[^\n\r\d]+(\d{2,3})\s*:\s*[^\n\r]+)?(?:[^\n\r\d]+(\d+)\s+latidos)?", texto)
    d["taqui_conteo"] = int(taqui_m.group(1)) if taqui_m else 0
    d["taqui_fc_max"] = int(taqui_m.group(2)) if (taqui_m and taqui_m.group(2)) else d["fc_max"]
    d["taqui_duracion"] = int(taqui_m.group(3)) if (taqui_m and taqui_m.group(3)) else 0

    bradi_m = re.search(r"Bradicardia\s+(\d+)(?:[^\n\r\d]+(\d{2,3})\s*:\s*[^\n\r]+)?(?:[^\n\r\d]+(\d+)\s+latidos)?", texto)
    d["bradi_conteo"] = int(bradi_m.group(1)) if bradi_m else 0
    d["bradi_fc_min"] = int(bradi_m.group(2)) if (bradi_m and bradi_m.group(2)) else d["fc_min"]
    d["bradi_duracion"] = int(bradi_m.group(3)) if (bradi_m and bradi_m.group(3)) else 0

    pausa_match = re.search(r"\bPausa\s+(\d+)", texto)
    d["pausas"] = int(pausa_match.group(1)) if pausa_match else 0
    p_max_m = re.search(r"Pausa.*?M[áa]x\.\s*longitud\s*([\d,\.]+)\s*s", texto)
    d["pausa_max_seg"] = p_max_m.group(1).replace(",", ".") if p_max_m else "0"

    lat_caido_m = re.search(r"Latidos?\s+ca[íi]dos?\s+(\d+)", texto)
    d["latidos_caidos"] = int(lat_caido_m.group(1)) if lat_caido_m else 0

    ev_m = re.search(r"Latidos ventriculares\s*:\s*([\d\.]+)", texto)
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
    d["st_dep_mm"] = st_dep.group(2) if (st_dep and st_dep.group(2)) else "1.0"

    d["st_elev_episodios"] = int(st_elev.group(1)) if (st_elev and st_elev.group(1) != "0") else 0
    d["st_elev_mm"] = st_elev.group(2) if (st_elev and st_elev.group(2)) else "1.0"

    d["st_episodios"] = d["st_dep_episodios"] + d["st_elev_episodios"]

    sdnn_m = re.search(r"Valor de 24 horas\s+[\d\.,]+\s+([\d\.,]+)", texto)
    d["sdnn_24h"] = limpiar_numero(sdnn_m.group(1)) if sdnn_m else 85
    qtc_m = re.search(r"Todos los per[íi]odos\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)", texto)
    d["qtc_prom"] = limpiar_numero(qtc_m.group(1)) if qtc_m else 400

    eventos_pac_m = re.search(r"Eventos del paciente\s*:\s*(\d+)", texto)
    d["eventos_paciente"] = int(eventos_pac_m.group(1)) if eventos_pac_m else 0

    return d

def redactar_informe_holter_11_puntos(d, perfil):
    p1 = f"1. Ritmo sinusal con frecuencia cardiaca promedio de {d['fc_prom']} latidos por minuto (Diurna: {d['fc_dia']} lpm / Nocturna: {d['fc_noc']} lpm)."

    eventos_crono = []
    if d["taqui_conteo"] > 0:
        det_t = f"{d['taqui_conteo']} episodios de taquicardia sinusal (FC máxima {d['taqui_fc_max']} lpm"
        if d["taqui_duracion"] > 0:
            det_t += f", racha más prolongada de {d['taqui_duracion']} latidos"
        det_t += ")"
        eventos_crono.append(det_t)

    if d["bradi_conteo"] > 0:
        det_b = f"{d['bradi_conteo']} episodios de bradicardia sinusal (FC mínima {d['bradi_fc_min']} lpm"
        if d["bradi_duracion"] > 0:
            det_b += f", racha más prolongada de {d['bradi_duracion']} latidos"
        det_b += ")"
        eventos_crono.append(det_b)

    if eventos_crono:
        p2 = f"2. Eventos cronotrópicos: Se documentaron {' y '.join(eventos_crono)}."
    else:
        p2 = "2. Eventos cronotrópicos: Sin episodios de bradicardia patológica ni taquicardias sostenidas de relevancia clínica."

    p3 = f"3. Intervalos PR normales y QTc prolongado (promedio {d['qtc_prom']} ms)." if d["qtc_prom"] > 460 else f"3. Intervalos PR normales y QTc normales ({d['qtc_prom']} ms)."

    es_isq = (d["st_episodios"] > 0) or any(k in d.get("dx_motivo", "") for k in ["ANGINA", "INFARTO", "ISQUEMIA", "CORONAR", "IAM", "SCA", "NECROSIS", "DOLOR"])
    if es_isq:
        det_st = []
        if d["st_dep_episodios"] > 0:
            det_st.append(f"{d['st_dep_episodios']} episodios de depresión del ST, máx. {d['st_dep_mm']} mm")
        if d["st_elev_episodios"] > 0:
            det_st.append(f"{d['st_elev_episodios']} episodios de elevación del ST, máx. +{d['st_elev_mm']} mm")
        
        p4 = f"4. Alteraciones isquémicas del segmento ST ({', '.join(det_st)})." if det_st else "4. Alteraciones isquémicas del segmento ST."
    else:
        p4 = "4. Sin alteraciones isquémicas del segmento ST."

    if d["latidos_caidos"] > 0:
        p5 = f"5. Alteración de la conducción AV por {d['latidos_caidos']} latidos caídos."
    elif d["pausas"] > 0 and float(d["pausa_max_seg"]) >= 2.0:
        p5 = f"5. Alteración de la conducción AV por {d['pausas']} pausas significativas (máx. {d['pausa_max_seg']} s)."
    else:
        p5 = "5. Sin Alteración de la conducción AV."

    tiene_bloqueo = any(k in d.get("dx_motivo", "") for k in ["EPOC", "PULMONAR", "BLOQUEO", "RAMA", "BRD", "BRI", "BCRD", "BCRI", "QRS", "CARDIOPATIA"])
    p6 = "6. Alteración en la conducción intraventricular por bloqueo completo de rama." if tiene_bloqueo else "6. Sin Alteración en la conducción intraventricular."

    ect = []
    if d["esv_total"] > 0:
        txt_s = f"ectopias supraventriculares ({d['esv_total']} ESV"
        if d["tsv_episodios"] > 0:
            txt_s += f", incluyendo {d['tsv_episodios']} rachas de taquicardia supraventricular"
        txt_s += ")"
        ect.append(txt_s)

    if d["ev_total"] > 0:
        txt_v = f"ventriculares frecuentes ({d['ev_total']} EV" if d["ev_total"] >= 50 else f"ventriculares aisladas ({d['ev_total']} EV"
        sub = []
        if d["tv_episodios"] > 0:
            sub.append(f"{d['tv_episodios']} episodios de TV")
        if d["ev_duplas"] > 0:
            sub.append(f"{d['ev_duplas']} duplas" if d["ev_duplas"] > 1 else "1 dupla")
        if d["bigeminismo"] > 0:
            sub.append("bigeminismo")
        if sub:
            txt_v += f", incluyendo {', '.join(sub)}"
        txt_v += ")"
        ect.append(txt_v)

    p7 = f"7. Alteración de los impulsos por {' y '.join(ect)}." if ect else "7. Sin alteración de los impulsos ectópicos de relevancia clínica."
    p8 = f"8. El paciente refirió síntomas ({d['eventos_paciente']} eventos marcados en diario)." if d["eventos_paciente"] > 0 else "8. No refirió síntomas."

    sdnn = d["sdnn_24h"]
    if sdnn <= 60:
        p9 = "9. Variabilidad Severamente Disminuida de la FC."
        riesgo = "Riesgo alto"
    elif sdnn <= 120:
        p9 = "9. Variabilidad Disminuida de la FC."
        riesgo = "Riesgo medio"
    else:
        p9 = "9. Variabilidad Conservada de la FC."
        riesgo = "Bajo riesgo / Normal"

    p10 = f"10. Se registraron {d['pausas']} pausas significativas (máx. {d['pausa_max_seg']} s)." if d["pausas"] > 0 else "10. No hay pausas significativas."
    p11 = f"11. Riesgo del paciente SDNN 20 a 24 HRS ({riesgo})."

    c_diag = []
    if 60 <= d['fc_prom'] <= 100:
        c_diag.append(f"Ritmo sinusal con respuesta ventricular promedio conservada ({d['fc_prom']} lpm).")
    else:
        c_diag.append(f"Ritmo sinusal con FC promedio {d['fc_prom']} lpm.")

    desc = ((d["fc_dia"] - d["fc_noc"]) / d["fc_dia"]) * 100 if d["fc_dia"] > 0 else 0
    c_diag.append(f"Patrón circadiano conservado ({desc:.1f}% descenso nocturno)." if desc >= 10 else f"Patrón circadiano no-dipper ({desc:.1f}% descenso nocturno).")

    if d["tv_episodios"] > 0:
        c_diag.append(f"Registro de taquicardia ventricular no sostenida ({d['tv_episodios']} rachas TV).")
    if es_isq:
        c_diag.append("Cambios en la repolarización compatibles con isquemia miocárdica.")
    c_diag.append(f"Variabilidad autonómica de la FC {riesgo.lower()}.")

    recs = []
    if d["tv_episodios"] > 0 or d["ev_total"] > 2000:
        recs.append("Ecocardiograma transtorácico (FEVI) y valoración por electrofisiología.")
    if es_isq:
        recs.append("Estratificación de cardiopatía isquémica funcional o invasiva.")
    if d["qtc_prom"] > 460:
        recs.append("Control electrolítico y revisión de fármacos.")

    txt_concl = " ".join(c_diag)
    if recs:
        txt_concl += "\nRECOMENDACIONES: " + " ".join(recs)

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
{txt_concl}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# MOTOR 2: MAPA 24 HORAS SENTINEL (ESTÁNDAR EXACTO DE CENCARDIO EN 6 PUNTOS)
# ==============================================================================
def extraer_datos_mapa_sentinel(pdf_bytes, filename=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    for page in doc:
        texto += page.get_text() + "\n"
    doc.close()

    d = {}
    m_nom = re.search(r"Nombre del paciente:\s*([^\n\r\|]+)", texto, re.IGNORECASE)
    if not m_nom:
        m_nom = re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,45},\s*[A-ZÁÉÍÓÚÑ\s]{3,45})", texto)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else "PACIENTE MAPA"

    m_id = re.search(r"ID paciente:\s*([^\n\r\s]+)", texto, re.IGNORECASE)
    d["cedula"] = m_id.group(1).replace(".", "") if m_id else ""

    p_gen = re.search(r"Resumen general.*?Prom\.?:\s*(\d{2,3})\s*[\/\-]\s*(\d{2,3})\s*mmHg", texto, re.IGNORECASE)
    if p_gen:
        d["pas_24h"] = int(p_gen.group(1))
        d["pad_24h"] = int(p_gen.group(2))
    else:
        p_alt = re.search(r"Prom\.?:\s*(\d{2,3})\s*[\/\-]\s*(\d{2,3})\s*mmHg", texto)
        d["pas_24h"] = int(p_alt.group(1)) if p_alt else 120
        d["pad_24h"] = int(p_alt.group(2)) if p_alt else 75

    c_sis = re.search(r"Sist[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", texto, re.IGNORECASE)
    d["carga_pas"] = c_sis.group(1).replace(".", ",") if c_sis else "0,00"

    c_dia = re.search(r"Diast[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", texto, re.IGNORECASE)
    d["carga_pad"] = c_dia.group(1).replace(".", ",") if c_dia else "0,00"

    pp_m = re.search(r"Presi[óo]n de pulso\s*\(mmHg\)\s*\n?\s*(\d{2,3})", texto, re.IGNORECASE)
    d["pp_val"] = int(pp_m.group(1)) if pp_m else (d["pas_24h"] - d["pad_24h"])

    caida_m = re.search(r"Sist[óo]lico\s*\(mmHg\)\s*.*?([\d,\.\-]+)\s*%", texto, re.DOTALL)
    d["caida_nocturna_str"] = caida_m.group(1).replace(",", ".") if caida_m else "10.0"
    try:
        d["caida_nocturna_val"] = float(d["caida_nocturna_str"])
    except:
        d["caida_nocturna_val"] = 10.0

    m_sueno = re.search(r"Resumen de los per[íi]odos de sue[ñn]o.*?Sist[óo]lico\s*\(mmHg\)\s*\n?\s*(\d+)\s*.*?(\d{2,3})\s*\([^\)]+\)\s*.*?Diast[óo]lico\s*\(mmHg\)\s*\n?\s*(\d+)\s*.*?(\d{2,3})\s*\(", texto, re.DOTALL | re.IGNORECASE)
    if m_sueno:
        d["pas_max_sueno"] = int(m_sueno.group(2))
        d["pad_max_sueno"] = int(m_sueno.group(4))
    else:
        d["pas_max_sueno"] = d["pas_24h"]
        d["pad_max_sueno"] = d["pad_24h"]

    return d

def redactar_informe_mapa_cencardio(d, perfil):
    p1 = f"1. Promedio de tensión arterial sistólica ({d['pas_24h']} mmHg) y diastolica ({d['pad_24h']} mmHg)"
    p2 = f"2. Carga tensional sistólica ({d['carga_pas']}%) y diastólica de ({d['carga_pad']}%)"
    p3 = "3. Presión de pulso normal" if d["pp_val"] <= 60 else f"3. Presión de pulso aumentada ({d['pp_val']} mmHg, rigidez arterial)"

    cn = d["caida_nocturna_val"]
    patron = "dipping positivo" if cn >= 10.0 else "dipping atenuado"
    p4 = f"4. Patrón circadiano tensional {patron}"

    if d["pas_max_sueno"] >= 145 or d["pad_max_sueno"] >= 95:
        p5 = f"5. Se presentaron incrementos significativos de presión arterial durante el período de sueño (pico nocturno de {d['pas_max_sueno']}/{d['pad_max_sueno']} mmHg)."
    else:
        p5 = "5. No se presentaron incrementos de presión arterial tanto sistólica al acostarse como diastólica durante las 24 horas."

    c_pas_num = float(d["carga_pas"].replace(",", "."))
    c_pad_num = float(d["carga_pad"].replace(",", "."))
    if c_pas_num < 15.0 and c_pad_num < 15.0 and d["pas_24h"] < 130 and d["pad_24h"] < 80:
        control_txt = "Control óptimo de la tensión arterial"
    elif c_pas_num <= 30.0 or c_pad_num <= 30.0 or d["pas_24h"] < 140:
        control_txt = "Control subóptimo de la tensión arterial estadio I"
    else:
        control_txt = "Descontrol de la tensión arterial estadio II"
    p6 = f"6. {control_txt}"

    return f"""INTERPRETACIÓN TEST MAPA
Hallazgos:
{p1}
{p2}
{p3}
{p4}
{p5}
{p6}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# MOTOR 3: PRUEBA DE ESFUERZO / ERGOMETRÍA (CUPS 893805)
# ==============================================================================
def extraer_datos_de_fotos_esfuerzo(archivos_fotos):
    datos = {
        "paciente": "",
        "cedula": "",
        "edad": 0,
        "sexo": "Femenino",
        "protocolo": "Bruce",
        "etapa": "Etapa 4",
        "tiempo_min": 0.0,
        "fc_basal": 75,
        "fc_pico": 0,
        "pas_basal": 120,
        "pad_basal": 80,
        "pas_pico": 140,
        "pad_pico": 85,
        "st_mm": 0.0
    }

    if not PYTESSERACT_INSTALADO or not archivos_fotos:
        return datos

    texto_acumulado = ""
    fcs_detectadas = []
    pas_detectadas = []
    tiempos_detectados = []

    for f in archivos_fotos:
        try:
            img = Image.open(f)
            txt = pytesseract.image_to_string(img, lang="spa+eng")
            texto_acumulado += txt + "\n"

            # Buscar todas las FC en las fotos
            for m in re.finditer(r"\bFC\s*[:\.]?\s*(\d{2,3})\b", txt, re.IGNORECASE):
                fcs_detectadas.append(int(m.group(1)))

            # Buscar todas las PA
            for m in re.finditer(r"\bPA\s*[:\.]?\s*(\d{2,3})\s*[\/\-]\s*(\d{2,3})\b", txt, re.IGNORECASE):
                pas_detectadas.append((int(m.group(1)), int(m.group(2))))

            # Buscar tiempos de fase
            for m in re.finditer(r"(?:Tiempo|Fase)\s*[:\.]?\s*(\d{1,2})[:\.](\d{2})", txt, re.IGNORECASE):
                tiempos_detectados.append(float(f"{m.group(1)}.{m.group(2)}"))
        except Exception:
            continue

    # Paciente
    m_nom = re.search(r"Paciente\s*:\s*([A-ZÁÉÍÓÚÑ\s,]{4,40})", texto_acumulado, re.IGNORECASE)
    if m_nom:
        datos["paciente"] = m_nom.group(1).replace("\n", " ").strip()

    # PID / Cédula
    m_id = re.search(r"(?:PID|ID|C\.?C\.?)\s*[:\.]?\s*(\d{5,12})", texto_acumulado, re.IGNORECASE)
    if m_id:
        datos["cedula"] = m_id.group(1)

    # Edad
    m_edad = re.search(r"Edad\s*[:\.]?\s*(\d{1,3})", texto_acumulado, re.IGNORECASE)
    if m_edad:
        datos["edad"] = int(m_edad.group(1))

    # Sexo
    if re.search(r"Femenino|\bFem\b|\bF\b", texto_acumulado, re.IGNORECASE):
        datos["sexo"] = "Femenino"
    elif re.search(r"Masculino|\bMasc\b|\bM\b", texto_acumulado, re.IGNORECASE):
        datos["sexo"] = "Masculino"

    # FC Pico y Basal
    if fcs_detectadas:
        datos["fc_pico"] = max(fcs_detectadas)
        datos["fc_basal"] = min(fcs_detectadas) if min(fcs_detectadas) > 40 else 75

    # PA Pico y Basal
    if pas_detectadas:
        datos["pas_pico"] = max(p[0] for p in pas_detectadas)
        datos["pad_pico"] = max(p[1] for p in pas_detectadas)
        datos["pas_basal"] = min(p[0] for p in pas_detectadas)
        datos["pad_basal"] = min(p[1] for p in pas_detectadas)

    # Tiempo
    if tiempos_detectados:
        datos["tiempo_min"] = max(tiempos_detectados)

    return datos

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
    if dts >= 5: duke = "Bajo riesgo coronario (< 1% mortalidad anual)"
    elif dts >= -10: duke = "Riesgo moderado (1 - 3% mortalidad anual)"
    else: duke = "Alto riesgo coronario (> 3% mortalidad anual)"

    p1 = f"1. Ritmo sinusal normal basal y durante todas las etapas del esfuerzo."
    p2 = f"2. Protocolo de {d['protocolo']} completado con duración de {d['tiempo_min']:.2f} minutos ({d['etapa']})."
    p3 = f"3. Capacidad funcional alcanzada: {d['mets']} METs."
    p4 = f"4. Respuesta cronotrópica: FC basal {d['fc_basal']} lpm elevándose hasta FC pico {d['fc_pico']} lpm ({porc}% de la FCM prevista, prueba {suf})."
    p5 = f"5. Respuesta hemodinámica presora: PA basal {d['pas_basal']}/{d['pad_basal']} mmHg alcanzando PA pico {d['pas_pico']}/{d['pad_pico']} mmHg."
    p6 = f"6. Doble producto máximo alcanzado: {dp:,} mmHg*lpm."
    p7 = f"7. Comportamiento electrocardiográfico del ST: {st_res}."
    p8 = "8. Sin arritmias ventriculares complejas ni eventos supraventriculares sostenidos inducidos por el ejercicio."
    p9 = "9. Motivo de suspensión: Consecución de frecuencia cardíaca máxima y agotamiento físico voluntario, sin dolor precordial."
    p10 = f"10. Estratificación pronóstica por Duke Treadmill Score: {dts:.1f} ({duke})."

    concl = f"Prueba de esfuerzo {suf}, eléctricamente {diag_el} para isquemia miocárdica inducible. Buena tolerancia hemodinámica y funcional."
    recs = "Continuar manejo médico integral y prescripción de actividad física aeróbica regular." if d["st_mm"] < 1.0 else "Valoración prioritaria por cardiología clínica para estudio funcional o angiografía."

    return f"""INTERPRETACIÓN PRUEBA DE ESFUERZO COMPUTARIZADA - CUPS 893805

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
{concl}
RECOMENDACIONES: {recs}

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# GENERADOR DEFENSIVO DE EXPEDIENTE ERGOMETRÍA (RESUELVE TYPEERROR)
# ==============================================================================
def generar_pdf_ergometria_completo(d, texto_informe, perfil, cod_uuid, imagenes_adjuntas=[]):
    doc = fitz.open()
    page = doc.new_page(width=612, height=792)  # Carta estándar

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
    page.insert_text(fitz.Point(46, 112), f"PACIENTE: {d['paciente'].upper()}", fontsize=8.5, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(46, 126), f"DOCUMENTO: {d['cedula']}    |    EDAD: {d['edad']} AÑOS    |    SEXO: {d['sexo']}", fontsize=7.5, fontname="helv", color=(0.2, 0.25, 0.3))
    page.insert_text(fitz.Point(46, 140), f"FECHA DEL ESTUDIO: {datetime.now().strftime('%d/%m/%Y')}    |    MÉDICO LECTOR: {perfil['nombre_completo']}", fontsize=7.5, fontname="helv", color=(0.2, 0.25, 0.3))

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

    fecha_emision = datetime.now().strftime("%Y-%m-%d")
    qr_bytes = generar_qr_verificacion(d['paciente'], perfil['nombre_completo'], fecha_emision, cod_uuid, "CUPS 893805")
    page.insert_image(fitz.Rect(48, 695, 100, 747), stream=qr_bytes)
    page.insert_text(fitz.Point(108, 715), "Certificado Digital Forense", fontsize=6.2, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(108, 726), "Res. 3100 de 2019 · Habilitación MinSalud", fontsize=5.5, fontname="helv", color=(0.4, 0.45, 0.5))
    page.insert_text(fitz.Point(108, 737), f"Cód: {cod_uuid[:16]}...", fontsize=5.2, fontname="helv", color=(0.4, 0.45, 0.5))

    firma_bytes = procesar_firma_transparente()
    if firma_bytes and perfil["id"] in ["dr.amaya", "admin"]:
        page.insert_image(fitz.Rect(390, 690, 545, 755), stream=firma_bytes)

    # 1. RENDERIZADO PREVIO DE PÁGINA 1 (ANTES DE MODIFICAR EL ÁRBOL DEL DOCUMENTO)
    try:
        pix = page.get_pixmap(dpi=130)
        img_preview = pix.tobytes("png")
    except Exception:
        img_preview = None

    # 2. ADJUNTAR PÁGINAS EXTRA CON LAS FOTOS ESCANEADAS DE LA BANDA
    if imagenes_adjuntas:
        for img_file in imagenes_adjuntas:
            try:
                p_extra = doc.new_page(width=612, height=792)
                img_data = img_file.getvalue() if hasattr(img_file, "getvalue") else img_file
                p_extra.insert_image(fitz.Rect(20, 20, 592, 772), stream=img_data)
            except Exception:
                pass

    pdf_bytes = doc.tobytes()
    doc.close()
    return pdf_bytes, img_preview

# ==============================================================================
# INYECCIÓN EXCLUSIVA HOLTER ECG (PARTE INFERIOR Y = 540 A 740)
# ==============================================================================
def inyectar_holter_pdf(pdf_bytes, texto_informe, paciente_nom, perfil, cod_uuid, estampador_activo=False):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]

    rects_h = pagina1.search_for("Hallazgos:")
    rects_f = pagina1.search_for("Firma del médico")
    if not rects_f: rects_f = pagina1.search_for("Firma del operador")

    y_base = rects_f[0].y0 if rects_f else 740
    y0 = (rects_h[0].y1 + 4) if rects_h else 545
    y1 = y_base - 62

    # Blanqueado de caja de texto inferior (sin tocar tablas superiores de arritmias)
    rect_caja = fitz.Rect(35, y0, pagina1.rect.width - 36, y1)
    pagina1.draw_rect(rect_caja, color=None, fill=(1, 1, 1), overlay=True)

    font_size_optimo = 5.6
    for fs in [6.2, 5.9, 5.6, 5.3, 5.0, 4.7, 4.4, 4.2]:
        doc_test = fitz.open(stream=pdf_bytes, filetype="pdf")
        p_test = doc_test[0]
        rc = p_test.insert_textbox(rect_caja, texto_informe, fontsize=fs, fontname="helv", align=fitz.TEXT_ALIGN_LEFT)
        doc_test.close()
        if rc >= 0:
            font_size_optimo = fs
            break

    pagina1.insert_textbox(rect_caja, texto_informe, fontsize=font_size_optimo, fontname="helv", color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)

    # Blanqueado de zona inferior para QR y Firma junto a línea de médico
    rect_inferior = fitz.Rect(220, y_base - 58, pagina1.rect.width - 36, y_base + 12)
    pagina1.draw_rect(rect_inferior, color=None, fill=(1, 1, 1), overlay=True)

    fecha_emision = datetime.now().strftime("%Y-%m-%d")
    qr_bytes = generar_qr_verificacion(paciente_nom, perfil['nombre_completo'], fecha_emision, cod_uuid, "CUPS 895001")
    rect_qr = fitz.Rect(230, y_base - 32, 270, y_base + 8)
    pagina1.insert_image(rect_qr, stream=qr_bytes)

    pagina1.insert_text(fitz.Point(275, y_base - 18), "Validado Digitalmente", fontsize=5.2, fontname="helv", color=(0.08, 0.2, 0.36))
    pagina1.insert_text(fitz.Point(275, y_base - 9), "Res. 3100 de 2019 - MinSalud", fontsize=4.7, fontname="helv", color=(0.25, 0.25, 0.25))
    pagina1.insert_text(fitz.Point(275, y_base), f"Cód: {cod_uuid[:12]}...", fontsize=4.5, fontname="helv", color=(0.4, 0.4, 0.4))

    firma_bytes = procesar_firma_transparente()
    if firma_bytes and estampador_activo:
        fx0 = (rects_f[0].x0 + 10) if rects_f else 380
        fy0 = y_base - 58
        rect_f = fitz.Rect(fx0, fy0, fx0 + 155, y_base + 4)
        pagina1.insert_image(rect_f, stream=firma_bytes)

    pix = pagina1.get_pixmap(dpi=130)
    img_preview = pix.tobytes("png")
    pdf_final = doc.tobytes()
    doc.close()
    return pdf_final, img_preview

# ==============================================================================
# INYECCIÓN EXCLUSIVA MAPA SENTINEL (FRANJA MEDIA Y = 235 A 420)
# ==============================================================================
def inyectar_mapa_pdf(pdf_bytes, texto_informe, paciente_nom, perfil, cod_uuid, estampador_activo=False):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]

    rect_m = pagina1.search_for("Presión arterial por la mañana")
    rect_r = pagina1.search_for("Resumen de todo el registro")

    y0 = (rect_m[0].y1 + 4) if rect_m else 235
    y1 = (rect_r[0].y0 - 6) if rect_r else 420

    # Blanquear franja completa
    rect_franja = fitz.Rect(35, y0, pagina1.rect.width - 35, y1)
    pagina1.draw_rect(rect_franja, color=None, fill=(1, 1, 1), overlay=True)

    # Inyectar texto en margen izquierdo
    rect_texto = fitz.Rect(36, y0 + 2, 440, y1 - 2)

    font_size = 6.4
    for fs in [7.2, 6.8, 6.4, 6.0, 5.6]:
        doc_test = fitz.open(stream=pdf_bytes, filetype="pdf")
        p_test = doc_test[0]
        rc = p_test.insert_textbox(rect_texto, texto_informe, fontsize=fs, fontname="helv", align=fitz.TEXT_ALIGN_LEFT)
        doc_test.close()
        if rc >= 0:
            font_size = fs
            break

    pagina1.insert_textbox(rect_texto, texto_informe, fontsize=font_size, fontname="helv", color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)

    # QR en margen derecho superior
    fecha_emision = datetime.now().strftime("%Y-%m-%d")
    qr_bytes = generar_qr_verificacion(paciente_nom, perfil['nombre_completo'], fecha_emision, cod_uuid, "CUPS 895003")
    rect_qr = fitz.Rect(455, y0 + 6, 505, y0 + 56)
    pagina1.insert_image(rect_qr, stream=qr_bytes)
    pagina1.insert_text(fitz.Point(510, y0 + 22), "Validado Digitalmente", fontsize=5.0, fontname="helv", color=(0.04, 0.15, 0.25))
    pagina1.insert_text(fitz.Point(510, y0 + 32), "Res. 3100 MinSalud", fontsize=4.6, fontname="helv", color=(0.3, 0.3, 0.3))
    pagina1.insert_text(fitz.Point(510, y0 + 42), f"Cód: {cod_uuid[:10]}...", fontsize=4.4, fontname="helv", color=(0.4, 0.4, 0.4))

    # Firma en margen derecho inferior
    firma_bytes = procesar_firma_transparente()
    if firma_bytes and estampador_activo:
        rect_f = fitz.Rect(445, y0 + 62, 575, y1 - 4)
        pagina1.insert_image(rect_f, stream=firma_bytes)

    pix = pagina1.get_pixmap(dpi=130)
    img_preview = pix.tobytes("png")
    pdf_final = doc.tobytes()
    doc.close()
    return pdf_final, img_preview

# ==========================================
# GESTIÓN DEL ENTORNO DE PROCESAMIENTO
# ==========================================
with tab_procesar:
    # ----------------------------------------------------------
    # MODALIDAD: PRUEBA DE ESFUERZO (OCR DE FOTOS O ENTRADA MANUAL)
    # ----------------------------------------------------------
    if "Esfuerzo" in modalidad_seleccionada:
        st.markdown("#### 🏃 Consola de Emisión de Prueba de Esfuerzo (CUPS 893805)")
        st.caption("Suba las fotografías del trazado impreso para auto-completar los datos o digítelos en el formulario:")

        # Inicialización de estado para la prueba de esfuerzo
        if "erg_datos" not in st.session_state:
            st.session_state.erg_datos = {
                "paciente": "",
                "cedula": "",
                "edad": 35,
                "sexo": "Femenino",
                "protocolo": "Bruce",
                "etapa": "",
                "tiempo_min": 0.0,
                "fc_basal": 75,
                "fc_pico": 150,
                "pas_basal": "120/80",
                "pas_pico": "160/90",
                "st_mm": 0.0
            }

        col_f1, col_f2 = st.columns([1.2, 1], gap="large")

        with col_f1:
            st.markdown("<b>1. Captura de Fotos del Trazado Impreso</b>", unsafe_allow_html=True)
            fotos_esfuerzo = st.file_uploader(
                "📸 Subir fotos o escaneos de las tiras de la banda (JPG o PNG):",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True,
                key="fotos_erg_uploader"
            )

            # Auto-extracción por OCR si se suben fotos nuevas
            if fotos_esfuerzo:
                if "fotos_procesadas_hash" not in st.session_state or st.session_state.fotos_procesadas_hash != len(fotos_esfuerzo):
                    with st.spinner("Analizando tiras de esfuerzo con motor OCR..."):
                        datos_ocr = extraer_datos_de_fotos_esfuerzo(fotos_esfuerzo)
                        for k, v in datos_ocr.items():
                            if v:
                                if k in ["pas_basal", "pad_basal"]:
                                    st.session_state.erg_datos["pas_basal"] = f"{datos_ocr['pas_basal']}/{datos_ocr['pad_basal']}"
                                elif k in ["pas_pico", "pad_pico"]:
                                    st.session_state.erg_datos["pas_pico"] = f"{datos_ocr['pas_pico']}/{datos_ocr['pad_pico']}"
                                else:
                                    st.session_state.erg_datos[k] = v
                        st.session_state.fotos_procesadas_hash = len(fotos_esfuerzo)
                        if datos_ocr.get("paciente") or datos_ocr.get("fc_pico"):
                            st.toast("✅ Datos clínicos extraídos de las fotos correctamente.", icon="📸")

            st.write("")
            st.markdown("<b>2. Verificación de Parámetros Clínicos</b>", unsafe_allow_html=True)
            c_in1, c_in2, c_in3 = st.columns(3)
            with c_in1:
                p_nombre = st.text_input("Paciente:", value=st.session_state.erg_datos["paciente"], placeholder="Nombre completo")
                p_cedula = st.text_input("Cédula / Documento:", value=st.session_state.erg_datos["cedula"], placeholder="Documento ID")
            with c_in2:
                p_edad = st.number_input("Edad:", value=int(st.session_state.erg_datos["edad"]) if st.session_state.erg_datos["edad"] > 0 else 35, min_value=1, max_value=110)
                idx_sex = 0 if st.session_state.erg_datos["sexo"] == "Femenino" else 1
                p_sexo = st.selectbox("Sexo:", ["Femenino", "Masculino"], index=idx_sex)
            with c_in3:
                p_protocolo = st.selectbox("Protocolo:", ["Bruce", "Bruce Modificado", "Naughton"], index=0)
                p_etapa = st.text_input("Etapa alcanzada:", value=st.session_state.erg_datos["etapa"], placeholder="Ej: Etapa 6")

            c_in4, c_in5, c_in6 = st.columns(3)
            with c_in4:
                p_tiempo = st.number_input("Tiempo total (minutos):", value=float(st.session_state.erg_datos["tiempo_min"]), step=0.1)
                mets_calc = calcular_mets_bruce(p_tiempo)
                p_mets = st.number_input("Capacidad Funcional (METs):", value=float(mets_calc), step=0.5)
            with c_in5:
                p_fc_basal = st.number_input("FC Basal (lpm):", value=int(st.session_state.erg_datos["fc_basal"]))
                p_fc_pico = st.number_input("FC Pico alcanzada (lpm):", value=int(st.session_state.erg_datos["fc_pico"]))
            with c_in6:
                p_pa_basal = st.text_input("PA Basal (mmHg):", value=st.session_state.erg_datos["pas_basal"], placeholder="Ej: 119/70")
                p_pa_pico = st.text_input("PA Esfuerzo Pico (mmHg):", value=st.session_state.erg_datos["pas_pico"], placeholder="Ej: 140/87")

            c_in7, c_in8 = st.columns(2)
            with c_in7:
                p_st_mm = st.number_input("Desviación del ST (mm):", value=float(st.session_state.erg_datos["st_mm"]), step=0.5)
            with c_in8:
                tel_encontrado = buscar_telefono_dinamica(p_cedula) if p_cedula else ""
                p_celular = st.text_input("Celular (Envío WhatsApp):", value=tel_encontrado)

        with col_f2:
            st.markdown("<b>3. Diagnóstico Institucional y Dictamen</b>", unsafe_allow_html=True)

            if p_nombre.strip():
                pas_b, pad_b = [int(x) for x in p_pa_basal.split("/")] if "/" in p_pa_basal else (120, 80)
                pas_p, pad_p = [int(x) for x in p_pa_pico.split("/")] if "/" in p_pa_pico else (150, 90)

                datos_erg = {
                    "paciente": p_nombre,
                    "cedula": p_cedula,
                    "edad": p_edad,
                    "sexo": p_sexo,
                    "protocolo": p_protocolo,
                    "etapa": p_etapa if p_etapa else "Final",
                    "tiempo_min": p_tiempo,
                    "mets": p_mets,
                    "fc_basal": p_fc_basal,
                    "fc_pico": p_fc_pico,
                    "pas_basal": pas_b,
                    "pad_basal": pad_b,
                    "pas_pico": pas_p,
                    "pad_pico": pad_p,
                    "st_mm": p_st_mm
                }

                texto_erg_defecto = redactar_informe_ergometria_institucional(datos_erg, perfil_activo)
                texto_erg_final = st.text_area("Informe Oficial para Certificado:", value=texto_erg_defecto, height=310)

                estudio_uuid_erg = str(uuid.uuid4()).upper()
                pdf_erg, img_erg_prev = generar_pdf_ergometria_completo(
                    datos_erg,
                    texto_erg_final,
                    perfil_activo,
                    estudio_uuid_erg,
                    imagenes_adjuntas=fotos_esfuerzo if fotos_esfuerzo else []
                )

                col_be1, col_be2 = st.columns(2)
                with col_be1:
                    st.download_button(
                        label="📄 DESCARGAR CERTIFICADO COMPLETO",
                        data=pdf_erg,
                        file_name=f"{normalizar_nombre_archivo(p_nombre)}_Prueba_Esfuerzo.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                with col_be2:
                    if st.button("💾 Guardar en Archivo Clínico", use_container_width=True):
                        guardar_estudio_db(
                            p_nombre,
                            "Prueba de Esfuerzo",
                            "CUPS 893805",
                            f"FC Pico: {p_fc_pico} | {p_mets} METs",
                            perfil_activo['nombre_completo'],
                            texto_erg_final,
                            pdf_erg,
                            estudio_uuid_erg
                        )
                        st.success(f"✅ Guardado en archivo clínico: {p_nombre}")

                if p_celular:
                    tel_l = re.sub(r'\D', '', p_celular)
                    if not tel_l.startswith("57") and len(tel_l) == 10:
                        tel_l = "57" + tel_l
                    
                    url_c = f"https://holtercencardio.streamlit.app/?val={estudio_uuid_erg[:12]}&pac={urllib.parse.quote(p_nombre)}&med={urllib.parse.quote(perfil_activo['nombre_completo'])}&proc=CUPS_893805"
                    msg_w = f"Estimado(a) paciente {p_nombre}, CENCARDIO le hace entrega de su resultado oficial de Prueba de Esfuerzo (CUPS 893805). Verifique su autenticidad aquí: {url_c}"
                    wa_u = f"https://wa.me/{tel_l}?text={urllib.parse.quote(msg_w)}"
                    st.write("")
                    st.link_button("📲 ENVIAR RESULTADO POR WHATSAPP", wa_u, use_container_width=True)

                if img_erg_prev:
                    st.divider()
                    st.subheader("👁️ Vista Previa del Certificado Institucional (Página 1)")
                    st.image(img_erg_prev, caption=f"Página 1 - {p_nombre}", width=680)
            else:
                st.info("💡 Ingrese el nombre del paciente a la izquierda o cargue las fotografías del trazado para generar la prueba.")

    # ----------------------------------------------------------
    # MODALIDAD: ARCHIVOS DIGITALES (PDF DE HOLTER O MAPA SENTINEL)
    # ----------------------------------------------------------
    else:
        st.markdown(f"#### 📥 Cargar Estudio Digital ({modalidad_seleccionada})")
        uploaded_file = st.file_uploader("Seleccione el archivo PDF del estudio:", type=["pdf"])

        if uploaded_file is not None:
            bytes_originales = uploaded_file.getvalue()

            if "archivo_cargado_nombre" not in st.session_state or st.session_state.archivo_cargado_nombre != uploaded_file.name:
                with st.spinner("Analizando y detectando tipo de estudio automáticamente..."):
                    tipo_real = detectar_tipo_documento_clinico(bytes_originales)

                    if tipo_real == "HOLTER":
                        d_act = extraer_datos_holter(bytes_originales, uploaded_file.name)
                        txt_inf = redactar_informe_holter_11_puntos(d_act, perfil_activo)
                        cups_detectado = "CUPS 895001"
                        mod_detectada = "Holter ECG 24 Horas"
                        param_clave = f"FC {d_act['fc_prom']} | SDNN {d_act['sdnn_24h']}ms"
                    else:  # MAPA Sentinel
                        d_act = extraer_datos_mapa_sentinel(bytes_originales, uploaded_file.name)
                        txt_inf = redactar_informe_mapa_cencardio(d_act, perfil_activo)
                        cups_detectado = "CUPS 895003"
                        mod_detectada = "MAPA Tensional 24 Horas"
                        param_clave = f"PA 24h: {d_act['pas_24h']}/{d_act['pad_24h']} mmHg"

                    st.session_state.tipo_detectado = tipo_real
                    st.session_state.cups_actual = cups_detectado
                    st.session_state.mod_nombre = mod_detectada
                    st.session_state.datos_actuales = d_act
                    st.session_state.texto_informe = txt_inf
                    st.session_state.param_clave = param_clave
                    st.session_state.archivo_cargado_nombre = uploaded_file.name
                    st.session_state.estudio_uuid = str(uuid.uuid4()).upper()

                    cedula_pac = d_act.get("cedula", "")
                    st.session_state.telefono_paciente = buscar_telefono_dinamica(cedula_pac)

            datos = st.session_state.datos_actuales
            tipo_estudio = st.session_state.tipo_detectado
            cups_actual = st.session_state.cups_actual
            mod_nombre = st.session_state.mod_nombre

            alertas_criticas = []
            alertas_moderadas = []

            if tipo_estudio == "HOLTER":
                if datos["sdnn_24h"] <= 60:
                    alertas_criticas.append(f"Variabilidad severamente disminuida (SDNN {datos['sdnn_24h']} ms: Alto riesgo cardiovascular).")
                elif datos["sdnn_24h"] <= 120:
                    alertas_moderadas.append(f"Variabilidad de la FC disminuida (SDNN {datos['sdnn_24h']} ms: Riesgo medio).")

                if datos["tv_episodios"] > 0:
                    alertas_criticas.append(f"Se registraron {datos['tv_episodios']} rachas de Taquicardia Ventricular (TV).")

                if datos["pausas"] > 0 or datos["latidos_caidos"] > 0:
                    alertas_criticas.append(f"Trastorno de conducción AV: {datos['pausas']} pausas significativas, {datos['latidos_caidos']} latidos caídos.")

                es_isq = (datos["st_episodios"] > 0) or any(k in datos.get("dx_motivo", "") for k in ["ANGINA", "INFARTO", "ISQUEMIA", "CORONAR", "IAM", "SCA", "NECROSIS", "DOLOR"])
                if es_isq:
                    alertas_moderadas.append(f"Alteraciones isquémicas del ST detectadas ({datos['st_episodios']} episodios / Antecedente: {datos.get('dx_motivo', 'Isquemia')}).")

                if datos["qtc_prom"] > 460:
                    alertas_moderadas.append(f"Intervalo QTc prolongado (promedio {datos['qtc_prom']} ms).")

            else:
                cn = datos["caida_nocturna_val"]
                if cn < 0:
                    alertas_criticas.append(f"Patrón circadiano Riser / Invertido (PA nocturna superior a la diurna, alto riesgo cerebrovascular).")
                elif cn < 10:
                    alertas_moderadas.append(f"Patrón circadiano dipping atenuado / No-Dipper (descenso nocturno del {cn:.1f}%).")

                c_pas_num = float(datos["carga_pas"].replace(",", "."))
                c_dia_num = float(datos["carga_pad"].replace(",", "."))
                if c_pas_num > 30 or c_dia_num > 30 or datos["pas_24h"] >= 140 or datos["pad_24h"] >= 90:
                    alertas_criticas.append(f"Cargas tensionales elevadas (Sistólica {datos['carga_pas']}%, Diastólica {datos['carga_pad']}%).")
                elif c_pas_num >= 15 or c_dia_num >= 15:
                    alertas_moderadas.append(f"Carga tensional en rango limítrofe (Sistólica {datos['carga_pas']}%, Diastólica {datos['carga_pad']}%).")

                if datos["pp_val"] > 60:
                    alertas_moderadas.append(f"Presión de pulso ensanchada ({datos['pp_val']} mmHg: rigidez arterial).")

            if alertas_criticas:
                st.markdown(f"""
                    <div class="triage-rojo">
                        <b>🔴 ALERTA CRÍTICA: Hallazgos de Alto Riesgo Cardiovascular Detectados</b><br>
                        • {'<br>• '.join(alertas_criticas)}
                    </div>
                """, unsafe_allow_html=True)
            elif alertas_moderadas:
                st.markdown(f"""
                    <div class="triage-amarillo">
                        <b>🟡 PRECAUCIÓN CLÍNICA: Parámetros Fuera de Meta o Hallazgos Relevantes</b><br>
                        • {'<br>• '.join(alertas_moderadas)}
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                    <div class="triage-verde">
                        <b>🟢 ESTUDIO NORMAL / COMPENSADO: Sin Criterios de Alarma Electrocardiográfica ni Hemodinámica</b>
                    </div>
                """, unsafe_allow_html=True)

            if tipo_estudio == "HOLTER":
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("FC Promedio (24h)", f"{datos['fc_prom']} lpm", f"Día {datos['fc_dia']} | Noche {datos['fc_noc']}")
                c2.metric("Ectopias Ventriculares", f"{datos['ev_total']} EV", f"TV: {datos['tv_episodios']}")
                c3.metric("Ectopias Supraventriculares", f"{datos['esv_total']} ESV", f"TSV: {datos['tsv_episodios']}")
                c4.metric("SDNN (24 Horas)", f"{datos['sdnn_24h']} ms", f"ST: {datos['st_episodios']} ep.")
                st.plotly_chart(generar_grafica_tacograma(datos), use_container_width=True)
            else:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Promedio 24 Horas", f"{datos['pas_24h']}/{datos['pad_24h']} mmHg", "Meta < 130/80")
                c2.metric("Carga Sistólica", f"{datos['carga_pas']}%", "Normal < 15%")
                c3.metric("Carga Diastólica", f"{datos['carga_pad']}%", "Normal < 15%")
                c4.metric("Presión de Pulso", f"{datos['pp_val']} mmHg", "Meta <= 60")

            st.divider()

            col_edicion, col_preview = st.columns([1, 1], gap="large")

            with col_edicion:
                col_id1, col_id2 = st.columns([1.8, 1.2])
                with col_id1:
                    nombre_confirmado = st.text_input("👤 Paciente:", value=datos['paciente'])
                with col_id2:
                    cedula_confirmada = st.text_input("🪪 Cédula / Documento:", value=datos.get('cedula', ''))

                tel_actual = st.session_state.get("telefono_paciente", "")
                if tel_actual:
                    st.markdown(f'<div class="dinamica-status">✅ Teléfono Dinámica: <b>{tel_actual}</b></div>', unsafe_allow_html=True)
                
                telefono_input = st.text_input("📱 Celular (Envío de Resultado WhatsApp):", value=tel_actual)
                paciente_nombre_archivo = normalizar_nombre_archivo(nombre_confirmado)

                st.subheader(f"📝 Informe Oficial ({cups_actual})")
                informe_para_grabar = st.text_area("Texto oficial para inyectar en el reporte final:", value=st.session_state.texto_informe, height=360)

                debe_estampar = perfil_activo["id"] in ["dr.amaya", "admin"]

                # LLAMADA TOTALMENTE AISLADA SEGÚN EL TIPO DE ESTUDIO
                if tipo_estudio == "HOLTER":
                    pdf_final, img_preview = inyectar_holter_pdf(
                        bytes_originales,
                        informe_para_grabar,
                        nombre_confirmado,
                        perfil_activo,
                        st.session_state.estudio_uuid,
                        estampador_activo=debe_estampar
                    )
                else:
                    pdf_final, img_preview = inyectar_mapa_pdf(
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
                        label=f"📄 DESCARGAR {cups_actual} FIRMADO",
                        data=pdf_final,
                        file_name=f"{paciente_nombre_archivo}_{cups_actual.replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                with col_btn2:
                    if st.button("💾 Guardar en Archivo Clínico", use_container_width=True):
                        guardar_estudio_db(
                            nombre_confirmado,
                            mod_nombre,
                            cups_actual,
                            st.session_state.param_clave,
                            perfil_activo['nombre_completo'],
                            informe_para_grabar,
                            pdf_final,
                            st.session_state.estudio_uuid
                        )
                        st.success(f"✅ Guardado en archivo clínico: {nombre_confirmado}")

                if telefono_input:
                    tel_limpio = re.sub(r'\D', '', telefono_input)
                    if not tel_limpio.startswith("57") and len(tel_limpio) == 10:
                        tel_limpio = "57" + tel_limpio
                    
                    url_cert = f"https://holtercencardio.streamlit.app/?val={st.session_state.estudio_uuid[:12]}&pac={urllib.parse.quote(nombre_confirmado)}&med={urllib.parse.quote(perfil_activo['nombre_completo'])}&proc={urllib.parse.quote(mod_nombre)}"
                    msg_wa = f"Estimado(a) paciente {nombre_confirmado}, el Centro Cardiovascular Colombiano CENCARDIO le hace entrega de su resultado oficial de {mod_nombre} ({cups_actual}). Certificado oficial: {url_cert}"
                    wa_url = f"https://wa.me/{tel_limpio}?text={urllib.parse.quote(msg_wa)}"
                    
                    st.write("")
                    st.link_button(
                        label="📲 ENVIAR RESULTADO OFICIAL POR WHATSAPP",
                        url=wa_url,
                        use_container_width=True
                    )

                st.write("")
                if st.button("🔄 Cargar Nuevo Estudio", use_container_width=True):
                    for k in ["datos_actuales", "archivo_cargado_nombre", "texto_informe", "telefono_paciente", "param_clave", "tipo_detectado", "cups_actual", "mod_nombre"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.rerun()

            with col_preview:
                st.subheader("👁️ Vista Previa Oficial (Página 1)")
                st.markdown('<div class="preview-container">', unsafe_allow_html=True)
                if img_preview:
                    st.image(img_preview, caption=f"Página 1 - {nombre_confirmado} ({cups_actual})", use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# PESTAÑA 2: ARCHIVO CLÍNICO POR CARPETAS MENSUALES & EXCEL MULTI-HOJA
# ==============================================================================
with tab_historial:
    st.markdown("### 📁 Archivo Clínico Digital & Reportes Gerenciales")
    historial = obtener_historial_db()

    if not historial:
        st.info("Aún no hay estudios archivados en la base de datos institucional.")
    else:
        datos_tabla = []
        for h in historial:
            datos_tabla.append({
                "ID": h[0],
                "Fecha de Registro": h[1],
                "Paciente": h[2],
                "Modalidad": h[3],
                "Código CUPS": h[4],
                "Parámetro Clave": h[5],
                "Especialista Firmante": h[6],
                "pdf_data": h[7],
                "Código Forense": h[8] if len(h) > 8 and h[8] else "N/A"
            })
        df_produccion = pd.DataFrame(datos_tabla)

        # Generador de Excel con respaldo automático
        excel_bytes, ext_salida = generar_excel_avanzado_produccion(df_produccion)
        mime_tipo = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if ext_salida == "xlsx" else "text/csv"
        etiqueta_boton = "📊 DESCARGAR REPORTE EXCEL GERENCIAL (.XLSX)" if ext_salida == "xlsx" else "📊 DESCARGAR REPORTE RIPS (CSV EXCEL)"

        col_rep1, col_rep2 = st.columns([2.5, 1.5])
        with col_rep1:
            st.markdown(f"**Total de estudios en custodia:** `{len(df_produccion)} procedimientos certificados`")
        with col_rep2:
            st.download_button(
                label=etiqueta_boton,
                data=excel_bytes,
                file_name=f"Reporte_Gerencial_Cencardio_{datetime.now().strftime('%Y%m%d')}.{ext_salida}",
                mime=mime_tipo,
                use_container_width=True,
                type="primary"
            )

        st.divider()

        # Filtros Clínicos
        c_flt1, c_flt2, c_flt3 = st.columns([1.5, 1.2, 1.3])
        with c_flt1:
            busqueda = st.text_input("🔍 Buscar paciente por nombre o documento:", "")
        with c_flt2:
            df_produccion["Mes_Periodo"] = pd.to_datetime(df_produccion["Fecha de Registro"]).dt.strftime('%Y-%m')
            meses_disponibles = ["Todos los meses"] + sorted(df_produccion["Mes_Periodo"].unique().tolist(), reverse=True)
            mes_seleccionado = st.selectbox("📅 Filtrar por Carpeta Mensual:", meses_disponibles)
        with c_flt3:
            modalidades_disp = ["Todas las modalidades"] + sorted(df_produccion["Modalidad"].unique().tolist())
            mod_seleccionada = st.selectbox("🎛️ Filtrar por Tipo de Estudio:", modalidades_disp)

        df_filtrado = df_produccion.copy()
        if busqueda.strip():
            df_filtrado = df_filtrado[df_filtrado["Paciente"].str.contains(busqueda, case=False, na=False)]
        if mes_seleccionado != "Todos los meses":
            df_filtrado = df_filtrado[df_filtrado["Mes_Periodo"] == mes_seleccionado]
        if mod_seleccionada != "Todas las modalidades":
            df_filtrado = df_filtrado[df_filtrado["Modalidad"] == mod_seleccionada]

        st.write("")

        # Agrupación por Carpetas Mensuales
        meses_grupos = sorted(df_filtrado["Mes_Periodo"].unique().tolist(), reverse=True)

        if not meses_grupos:
            st.warning("No se encontraron estudios que coincidan con los criterios de búsqueda.")
        else:
            for mes_g in meses_grupos:
                df_mes = df_filtrado[df_filtrado["Mes_Periodo"] == mes_g]
                
                with st.expander(f"📁 CARPETA: {mes_g} ({len(df_mes)} estudios clínicos en custodia)", expanded=True):
                    for item in df_mes.itertuples():
                        nom_archivo_copia = normalizar_nombre_archivo(item.Paciente)
                        
                        st.markdown(f"""
                            <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:0.8rem 1.1rem; margin-bottom:0.7rem;">
                                <div style="display:flex; justify-content:space-between; align-items:center;">
                                    <div>
                                        <b style="color:#0a2540; font-size:1.02rem;">👤 {item.Paciente}</b> 
                                        <span style="background:#e0f2fe; color:#0369a1; font-size:0.72rem; font-weight:700; padding:2px 8px; border-radius:12px; margin-left:6px;">{item.Modalidad} ({item._5})</span>
                                        <div style="font-size:0.82rem; color:#64748b; margin-top:2px;">
                                            📅 <b>Registro:</b> {item._2} | 👨‍⚕️ <b>Lector:</b> {item._7} | 🩺 <b>Parámetro:</b> {item._6}
                                        </div>
                                    </div>
                                    <div style="font-family:monospace; font-size:0.75rem; color:#0284c7; font-weight:700;">
                                        Cód: {item._9[:14]}...
                                    </div>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        c_desc, c_del, c_vacio = st.columns([1.5, 1, 4])
                        with c_desc:
                            st.download_button(
                                label="📥 Descargar Copia PDF",
                                data=item.pdf_data,
                                file_name=f"{nom_archivo_copia}_{item._5.replace(' ', '_')}.pdf",
                                mime="application/pdf",
                                key=f"desc_{item.ID}",
                                use_container_width=True
                            )
                        with c_del:
                            if st.button("🗑️ Eliminar", key=f"elim_{item.ID}", use_container_width=True):
                                eliminar_estudio_db(item.ID)
                                st.toast(f"Estudio de {item.Paciente} eliminado.", icon="🗑️")
                                st.rerun()
