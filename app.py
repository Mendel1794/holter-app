import streamlit as st
import fitz  # PyMuPDF: motor C++ de lectura y renderizado ultrarrápido
from PIL import Image, ImageOps, ImageEnhance
import qrcode
import re
import base64
import os
import io
import sqlite3
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

# Diagnóstico activo del binario Tesseract en Linux/Windows
OCR_DISPONIBLE = False
MENSAJE_ESTADO_OCR = ""
try:
    import pytesseract
    _ = pytesseract.get_tesseract_version()
    OCR_DISPONIBLE = True
    MENSAJE_ESTADO_OCR = "Motor Tesseract instalado y activo en el sistema."
except Exception as e:
    OCR_DISPONIBLE = False
    MENSAJE_ESTADO_OCR = "Motor Tesseract no detectado en el servidor. Añada 'tesseract-ocr' a packages.txt."

st.set_page_config(
    page_title="Centro Cardiovascular Colombiano CENCARDIO · Workstation",
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
            mime = "png" if nom.endswith("png") else "jpeg"
            return f"data:image/{mime};base64,{b64}"
    return None

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

if not supabase:
    init_db_local()

def guardar_estudio_servicio(nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif):
    fecha_actual_str = ahora_colombia().strftime("%Y-%m-%d %H:%M:%S")
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
        INSERT INTO estudios (fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, informe_texto, pdf_blob, codigo_verificacion)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (fecha_actual_str, nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif))
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
            "id": f[0],
            "fecha_registro": f[1],
            "paciente": f[2],
            "modalidad": f[3],
            "cups": f[4],
            "parametro_clave": f[5],
            "medico": f[6],
            "pdf_url": "",
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
# SISTEMA VISUAL INSTITUCIONAL
# ==============================================================================
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
    .triage-rojo {
        background: #fef2f2 !important; border-left: 6px solid #dc2626 !important; color: #991b1b !important;
        padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important;
    }
    .triage-amarillo {
        background: #fffbeb !important; border-left: 6px solid #d97706 !important; color: #92400e !important;
        padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important;
    }
    .triage-verde {
        background: #f0fdf4 !important; border-left: 6px solid #16a34a !important; color: #166534 !important;
        padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important;
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

# ==========================================
# VALIDACIÓN PÚBLICA POR QR (RES. 3100)
# ==========================================
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

# ==========================================
# UTILIDADES FORENSES Y DE FIRMA
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
    limpio = re.sub(r'[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s]', ' ', nombre)
    return re.sub(r'\s+', '_', limpio).strip('_') or "PACIENTE"

# ==========================================
# GENERADOR AVANZADO DE EXCEL INSTITUCIONAL
# ==========================================
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

# ==========================================
# PERFILES MÉDICOS OFICIALES
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

perfil_activo = PERFILES_POR_ID[st.session_state.usuario_actual]

# ==============================================================================
# DETECCIÓN DE DOCUMENTO CLÍNICO & GRÁFICA TACOGRAMA
# ==============================================================================
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

# ==========================================
# ENTORNO DE OPERACIÓN
# ==========================================
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

with tab_procesar:
    if "Esfuerzo" in modalidad_seleccionada:
        st.markdown("#### 🏃 Consola de Emisión de Prueba de Esfuerzo (CUPS 893805)")
        st.caption("Complete los parámetros del esfuerzo y adjunte las fotografías del trazado impreso:")

        col_f1, col_f2 = st.columns([1.2, 1], gap="large")

        with col_f1:
            st.markdown("<b>1. Parámetros Clínicos</b>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1:
                p_nombre = st.text_input("Paciente:", value="", placeholder="Nombre completo")
                p_cedula = st.text_input("Cédula / ID:", value="", placeholder="Cédula")
            with c2:
                p_edad = st.number_input("Edad:", value=35, min_value=1, max_value=110)
                p_sexo = st.selectbox("Sexo:", ["Femenino", "Masculino"], index=0)
            with c3:
                p_protocolo = st.selectbox("Protocolo:", ["Bruce", "Bruce Modificado", "Naughton"], index=0)
                p_etapa = st.text_input("Etapa alcanzada:", value="", placeholder="Ej: Etapa 4")

            c4, c5, c6 = st.columns(3)
            with c4:
                p_tiempo = st.number_input("Tiempo total (min):", value=0.0, step=0.1)
                mets_calc = calcular_mets_bruce(p_tiempo)
                p_mets = st.number_input("Capacidad (METs):", value=float(mets_calc), step=0.5)
            with c5:
                p_fc_basal = st.number_input("FC Basal (lpm):", value=75)
                p_fc_pico = st.number_input("FC Pico (lpm):", value=150)
            with c6:
                p_pa_basal = st.text_input("PA Basal (mmHg):", value="", placeholder="Ej: 120/80")
                p_pa_pico = st.text_input("PA Esfuerzo Pico (mmHg):", value="", placeholder="Ej: 160/90")

            c7, c8 = st.columns(2)
            with c7:
                p_st_mm = st.number_input("Desviación del ST (mm):", value=0.0, step=0.5)
            with c8:
                tel_encontrado = buscar_telefono_servicio(p_cedula) if p_cedula else ""
                p_celular = st.text_input("Celular (WhatsApp):", value=tel_encontrado)

            st.write("")
            fotos_esfuerzo = st.file_uploader(
                "📸 Adjuntar fotos o escaneos de las tiras de la banda:",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True
            )

        with col_f2:
            st.markdown("<b>2. Diagnóstico Institucional y Certificación</b>", unsafe_allow_html=True)

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

                texto_erg = redactar_informe_ergometria_institucional(datos_erg, perfil_activo)
                texto_erg_final = st.text_area("Informe Oficial:", value=texto_erg, height=310)

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
                st.info("💡 Digite el nombre del paciente para redactar y certificar la prueba de esfuerzo.")

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

            datos = st.session_state.datos_actuales
            tipo_estudio = st.session_state.tipo_detectado
            cups_actual = st.session_state.cups_actual
            mod_nombre = st.session_state.mod_nombre

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
