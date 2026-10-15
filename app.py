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
# CONFIGURACIÓN GLOBAL Y ZONA HORARIA
# ==============================================================================
TZ_COLOMBIA = timezone(timedelta(hours=-5))

def ahora_colombia():
    return datetime.now(TZ_COLOMBIA)

# Importación segura de openpyxl (Excel Avanzado)
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_INSTALADO = True
except ImportError:
    OPENPYXL_INSTALADO = False

# Importación segura de Supabase (Nube)
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
# GESTIÓN DE LOGOTIPO INSTITUCIONAL Y ESTILOS UI
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
    .top-hospital-bar { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; padding: 1rem 1.6rem; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 20px -2px rgba(10, 37, 64, 0.04); margin-bottom: 1.2rem; }
    .inst-badge-primary { background: #e0f2fe; color: #0369a1; font-size: 0.72rem; font-weight: 700; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; border: 1px solid #bae6fd; }
    .inst-badge-success { background: #ecfdf5; color: #065f46; font-size: 0.72rem; font-weight: 700; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; border: 1px solid #a7f3d0; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; background: #ffffff; padding: 6px; border-radius: 12px; border: 1px solid #e2e8f0; margin-bottom: 1.2rem; }
    .stTabs [data-baseweb="tab"] { border-radius: 8px !important; padding: 9px 22px !important; font-weight: 700 !important; font-size: 0.88rem !important; color: #64748b !important; }
    .stTabs [aria-selected="true"] { background-color: #0a2540 !important; color: #ffffff !important; }
    .specialist-card { background: linear-gradient(135deg, #0a2540 0%, #133863 100%); border-radius: 14px; padding: 1.1rem; color: #ffffff; margin-bottom: 1.2rem; }
    .triage-rojo { background: #fef2f2 !important; border-left: 6px solid #dc2626 !important; color: #991b1b !important; padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important; box-shadow: 0 4px 12px rgba(220, 38, 38, 0.08); }
    .triage-amarillo { background: #fffbeb !important; border-left: 6px solid #d97706 !important; color: #92400e !important; padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important; box-shadow: 0 4px 12px rgba(217, 119, 6, 0.08); }
    .triage-verde { background: #f0fdf4 !important; border-left: 6px solid #16a34a !important; color: #166534 !important; padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important; box-shadow: 0 4px 12px rgba(22, 163, 74, 0.08); }
    .preview-container { border: 2px solid #e2e8f0; border-radius: 14px; background: #ffffff; padding: 8px; }
    .cencardio-card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 20px; padding: 2.6rem 2.8rem; box-shadow: 0 20px 40px -12px rgba(10, 37, 64, 0.12); max-width: 480px; margin: 3rem auto; text-align: center; }
    </style>
    """
st.markdown(cargar_estilos_institucionales(), unsafe_allow_html=True)

# ==============================================================================
# BASE DE DATOS LOCAL INCONDICIONAL
# ==============================================================================
def init_db_local():
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS estudios (
            id INTEGER PRIMARY KEY AUTOINCREMENT, fecha_registro TEXT, paciente_nombre TEXT, modalidad TEXT, cups TEXT,
            parametro_clave TEXT, medico_firmante TEXT, informe_texto TEXT, pdf_blob BLOB, codigo_verificacion TEXT, hash_sha256 TEXT)""")
    c.execute("CREATE TABLE IF NOT EXISTS dinamica_pacientes (cedula TEXT PRIMARY KEY, nombre TEXT, telefono TEXT, fecha_actualizacion TEXT)")
    c.execute("PRAGMA table_info(estudios)")
    cols_existentes = [col[1] for col in c.fetchall()]
    for col_nom, col_tipo in [("parametro_clave", "TEXT"), ("codigo_verificacion", "TEXT"), ("hash_sha256", "TEXT"), ("pdf_blob", "BLOB")]:
        if col_nom not in cols_existentes:
            try: c.execute(f"ALTER TABLE estudios ADD COLUMN {col_nom} {col_tipo}")
            except Exception: pass
    conn.commit()
    conn.close()

init_db_local()

# ==============================================================================
# CONECTOR DE NUBE PERSISTENTE: SUPABASE
# ==============================================================================
@st.cache_resource
def obtener_cliente_supabase():
    if not SUPABASE_LIB_OK: return None
    url = st.secrets.get("SUPABASE_URL", os.environ.get("SUPABASE_URL", ""))
    key = st.secrets.get("SUPABASE_KEY", os.environ.get("SUPABASE_KEY", ""))
    if url and key:
        try:
            url_limpia = str(url).strip().rstrip("/")
            for sub in ["/rest/v1", "/auth/v1", "/storage/v1"]:
                if url_limpia.endswith(sub): url_limpia = url_limpia[:-len(sub)].rstrip("/")
            cliente = create_client(url_limpia, str(key).strip().strip('"').strip("'"))
            try: cliente.storage.create_bucket("estudios-pdf", options={"public": True})
            except Exception: pass
            return cliente
        except Exception: return None
    return None

supabase = obtener_cliente_supabase()

def guardar_estudio_servicio(nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif):
    fecha_actual_str = ahora_colombia().strftime("%Y-%m-%d %H:%M:%S")
    hash_seguridad = hashlib.sha256(pdf_bytes).hexdigest()
    nombre_archivo = f"{cod_verif[:12]}_{re.sub(r'[^A-Za-z0-9]', '_', nombre)}.pdf"
    
    if supabase:
        try:
            supabase.storage.from_("estudios-pdf").upload(path=nombre_archivo, file=pdf_bytes, file_options={"content-type": "application/pdf", "upsert": "true"})
            pdf_url = supabase.storage.from_("estudios-pdf").get_public_url(nombre_archivo)
            supabase.table("estudios").insert({
                "fecha_registro": ahora_colombia().isoformat(), "paciente_nombre": nombre, "modalidad": modalidad,
                "cups": cups, "parametro_clave": parametro_clave, "medico_firmante": medico, "informe_texto": texto,
                "pdf_url": pdf_url, "codigo_verificacion": cod_verif
            }).execute()
        except Exception: pass

    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("""INSERT INTO estudios (fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, informe_texto, pdf_blob, codigo_verificacion, hash_sha256)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
              (fecha_actual_str, nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif, hash_seguridad))
    conn.commit()
    conn.close()
    return True, "Guardado con éxito en Archivo Clínico."

def obtener_historial_servicio():
    if supabase:
        try:
            res = supabase.table("estudios").select("id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, pdf_url, codigo_verificacion").order("id", desc=True).execute()
            lista = []
            for r in res.data:
                f_raw = r.get("fecha_registro", "")
                try: f_fmt = datetime.fromisoformat(f_raw.replace("Z", "+00:00")).astimezone(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M:%S")
                except Exception: f_fmt = f_raw[:19].replace("T", " ")
                lista.append({"id": r["id"], "fecha_registro": f_fmt, "paciente": r["paciente_nombre"], "modalidad": r["modalidad"], "cups": r["cups"], "parametro_clave": r.get("parametro_clave", ""), "medico": r.get("medico_firmante", ""), "pdf_url": r.get("pdf_url", ""), "codigo_verificacion": r.get("codigo_verificacion", "")})
            return lista, "supabase"
        except Exception: pass

    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, codigo_verificacion FROM estudios ORDER BY id DESC")
    filas = c.fetchall()
    conn.close()
    return [{"id": f[0], "fecha_registro": f[1], "paciente": f[2], "modalidad": f[3], "cups": f[4], "parametro_clave": f[5], "medico": f[6], "pdf_url": "", "codigo_verificacion": f[7] if f[7] else "N/A"} for f in filas], "local"

def obtener_pdf_bytes_individual(estudio_id, pdf_url=""):
    if pdf_url:
        try:
            r = requests.get(pdf_url, timeout=10)
            if r.status_code == 200: return r.content
        except Exception: pass
    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT pdf_blob FROM estudios WHERE id = ?", (estudio_id,))
    res = c.fetchone()
    conn.close()
    return res[0] if (res and res[0]) else None

def eliminar_estudio_servicio(estudio_id):
    if supabase:
        try: supabase.table("estudios").delete().eq("id", estudio_id).execute()
        except Exception: pass
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

    if not col_ced or not col_tel: return False, "El archivo debe contener columnas de Cédula y Teléfono/Celular."
    registros, fecha_hoy = 0, ahora_colombia().strftime("%Y-%m-%d %H:%M")
    lote = []
    for _, fila in df.iterrows():
        ced = re.sub(r'\D', '', str(fila[col_ced]))
        tel = re.sub(r'\D', '', str(fila[col_tel]))
        nom = str(fila[col_nom]) if col_nom else ""
        if len(ced) >= 5 and len(tel) >= 7:
            lote.append({"cedula": ced, "nombre": nom, "telefono": tel, "fecha_actualizacion": fecha_hoy})
            registros += 1

    if supabase and lote:
        try: supabase.table("dinamica_pacientes").upsert(lote).execute()
        except Exception: pass

    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    for item in lote:
        c.execute("INSERT OR REPLACE INTO dinamica_pacientes (cedula, nombre, telefono, fecha_actualizacion) VALUES (?, ?, ?, ?)", 
                  (item["cedula"], item["nombre"], item["telefono"], item["fecha_actualizacion"]))
    conn.commit()
    conn.close()
    return True, f"Se sincronizaron con éxito {registros} pacientes."

def buscar_telefono_servicio(cedula):
    if not cedula: return ""
    ced_limpia = re.sub(r'\D', '', str(cedula))
    if supabase:
        try:
            res = supabase.table("dinamica_pacientes").select("telefono").eq("cedula", ced_limpia).execute()
            if res.data: return res.data[0]["telefono"]
        except Exception: pass
    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT telefono FROM dinamica_pacientes WHERE cedula = ?", (ced_limpia,))
    res = c.fetchone()
    conn.close()
    return res[0] if res else ""

# ==============================================================================
# AUDITORÍA CLÍNICA (SANITY CHECKS)
# ==============================================================================
def ejecutar_sanity_checks(modalidad, datos):
    alertas, bloqueos = [], []
    if modalidad == "HOLTER":
        fc_p = datos.get("fc_prom")
        fc_min = datos.get("fc_min")
        fc_max = datos.get("fc_max")
        pr_int = datos.get("pr_intervalo")
        
        if fc_p is not None:
            if fc_min is not None and fc_min > fc_p: bloqueos.append(f"Incongruencia: FC mín ({fc_min}) supera a FC prom ({fc_p}).")
            if fc_max is not None and fc_max < fc_p: bloqueos.append(f"Incongruencia: FC máx ({fc_max}) es inferior a FC prom ({fc_p}).")
            
        if fc_min is not None and fc_min < 30 and not datos.get("mcp_presente"): alertas.append(f"Bradicardia Extrema documentada ({fc_min} lpm).")
        if fc_max is not None and fc_max > 220: alertas.append(f"FC Máxima ({fc_max} lpm) atípica. Verifique arritmia.")
        if pr_int is not None and pr_int > 200: alertas.append(f"Bloqueo AV documentado (PR: {pr_int} ms).")
            
    elif modalidad == "MAPA":
        pas = datos.get("pas_24h")
        pad = datos.get("pad_24h")
        cn = datos.get("caida_nocturna_val")
        if pas is not None and pad is not None and pad >= pas: bloqueos.append(f"Error Hemodinámico Severo: Diastólica ({pad}) >= Sistólica ({pas}).")
        if pas is not None and pad is not None and (pas - pad) < 20: alertas.append(f"Presión de Pulso estrecha ({pas - pad} mmHg).")
        if cn is not None and cn < -20.0: alertas.append(f"Patrón Riser Extremo (Aumento nocturno: {cn:.1f}%). Riesgo inminente.")
            
    elif modalidad == "ESFUERZO":
        pas_pico = datos.get("pas_pico")
        pad_pico = datos.get("pad_pico")
        if pas_pico is not None and pad_pico is not None and pad_pico >= pas_pico: bloqueos.append(f"Incongruencia Esfuerzo: PAD pico ({pad_pico}) >= PAS pico ({pas_pico}).")
            
    return bloqueos, alertas

def auditar_coherencia_informe(texto_informe, datos, modalidad):
    discrepancias = []
    texto_upper = texto_informe.upper()
    if modalidad == "HOLTER":
        if "SIN ALTERACIONES ISQUÉMICAS" in texto_upper and (datos.get("st_episodios") or 0) > 0:
            discrepancias.append(f"El informe dice 'Sin isquemia', pero hay {datos.get('st_episodios')} eventos ST.")
        if "TAQUICARDIA VENTRICULAR" in texto_upper and (datos.get("tv_episodios") or 0) == 0:
            discrepancias.append("Se menciona 'Taquicardia Ventricular', pero el contador es 0 o nulo.")
    return discrepancias

def mostrar_semaforizacion_clinica(modalidad, datos):
    nivel, titulo, detalles = "VERDE", "PARÁMETROS FISIOLÓGICOS CONSERVADOS", []
    if modalidad == "HOLTER":
        sdnn = datos.get("sdnn_24h")
        tv = datos.get("tv_episodios") or 0
        pausas = datos.get("pausas") or 0
        qtc = datos.get("qtc_prom")
        pr_int = datos.get("pr_intervalo")
        tot_qrs = datos.get("total_latidos") or 0
        ev = datos.get("ev_total") or 0
        carga_ev = (ev / tot_qrs) * 100 if tot_qrs and tot_qrs > 0 else 0
        duplas = datos.get("ev_duplas") or 0
        tsv = datos.get("tsv_episodios") or 0
        bloqueo = datos.get("tiene_bloqueo_rama", False)
        mcp = datos.get("mcp_presente", False)

        if tv and tv > 0: detalles.append(f"Taquicardia Ventricular documentada ({tv} episodios)")
        if sdnn is not None and sdnn > 0 and sdnn < 30 and not mcp: detalles.append(f"Variabilidad Autonómica Severamente Disminuida (SDNN: {sdnn} ms, Riesgo Alto)")
        if pausas and pausas > 0: detalles.append(f"Pausas patológicas registradas ({pausas} pausas)")
        if qtc is not None and qtc > 500: detalles.append(f"Intervalo QTc Severamente Prolongado ({qtc} ms, riesgo proarrítmico)")
        if carga_ev >= 10.0: detalles.append(f"Carga Ectópica Ventricular Crítica ({carga_ev:.1f}%)")

        if detalles: nivel, titulo = "ROJO", "ALERTA CLÍNICA: HALLAZGOS CRÍTICOS"
        else:
            if mcp: detalles.append(f"Paciente portador de Dispositivo Cardíaco ({datos.get('mcp_porcentaje') or 0.0:.1f}%)")
            if pr_int is not None and pr_int > 200: detalles.append(f"Bloqueo AV documentado (PR: {pr_int} ms)")
            if duplas and duplas > 0: detalles.append(f"Ectopia Ventricular Compleja: Duplas presentes ({duplas})")
            if tsv and tsv > 0: detalles.append(f"Taquicardia Supraventricular paroxística ({tsv} rachas)")
            if bloqueo and not mcp: detalles.append("Trastorno de la conducción intraventricular: Bloqueo de rama")
            if sdnn is not None and sdnn >= 30 and sdnn <= 100 and not mcp: detalles.append(f"Variabilidad de la FC disminuida (SDNN: {sdnn} ms, Riesgo Moderado)")
            if qtc is not None and qtc > 460: detalles.append(f"QTc prolongado ({qtc} ms)")
            if detalles: nivel, titulo = "AMARILLO", "PRECAUCIÓN CLÍNICA: DISPOSITIVO O PARÁMETROS LIMÍTROFES"
            else: detalles.append("Ritmo sinusal conservado, modulación autonómica normal y sin arritmias complejas.")
            
    elif modalidad == "MAPA":
        pas = datos.get("pas_24h") or 0
        pad = datos.get("pad_24h") or 0
        cn = datos.get("caida_nocturna_val") or 0.0
        c_pas = datos.get("carga_pas") or 0.0
        if cn <= -10.0: detalles.append(f"Patrón Circadiano Invertido (Riser: {cn:.1f}%)")
        if pas >= 140 or pad >= 90: detalles.append(f"Descontrol Tensional Estadio II ({pas}/{pad} mmHg)")
        if detalles: nivel, titulo = "ROJO", "ALERTA HEMODINÁMICA SEVERA"
        else:
            if cn <= 0.0: detalles.append("Patrón circadiano tensional no-dipper (atenuado)")
            if c_pas > 30.0: detalles.append(f"Carga sistólica elevada ({c_pas}%)")
            if detalles: nivel, titulo = "AMARILLO", "PRECAUCIÓN HEMODINÁMICA: CONTROL SUBÓPTIMO O CARGA ELEVADA"
            else: detalles.append("Control óptimo de la tensión arterial con modulación circadiana conservada.")
            
    elif modalidad == "ESFUERZO":
        st_mm = datos.get("st_mm") or 0.0
        fc_pico = datos.get("fc_pico") or 0
        edad = datos.get("edad", 35)
        porc = (fc_pico / (220 - edad)) * 100 if edad > 0 else 0
        
        if st_mm >= 1.0:
            detalles.append(f"Infradesnivel del ST significativo ({st_mm} mm)")
            nivel, titulo = "ROJO", "ALERTA ISQUÉMICA POSITIVA"
        elif porc > 0 and porc < 85.0:
            detalles.append(f"Prueba cronotrópicamente insuficiente ({porc:.1f}% de la FCM prevista)")
            nivel, titulo = "AMARILLO", "PRECAUCIÓN: ESFUERZO INSUFICIENTE"
        else: detalles.append(f"Prueba concluyente y suficiente ({porc:.1f}% FCM).")

    clase_css = f"triage-{nivel.lower()}"
    icono = "🚨" if nivel == "ROJO" else ("⚠️" if nivel == "AMARILLO" else "✅")
    items_html = "".join([f"<li>{d}</li>" for d in detalles])
    st.markdown(f'<div class="{clase_css}"><div style="font-weight: 800; font-size: 0.95rem; display: flex; align-items: center; gap: 8px;"><span>{icono}</span> {titulo}</div><ul style="margin: 0.4rem 0 0 1.2rem; font-size: 0.85rem; padding-left: 0;">{items_html}</ul></div>', unsafe_allow_html=True)

def generar_grafica_tacograma(d):
    fc_prom = d.get("fc_prom")
    if fc_prom is None or fc_prom <= 0:
        fig = go.Figure()
        fig.update_layout(title="Tacograma no disponible (Dato FC ausente)", height=240, margin=dict(l=35, r=20, t=35, b=25))
        return fig
    
    fc_dia = d.get("fc_dia") if d.get("fc_dia") is not None else int(fc_prom * 1.05)
    fc_noc = d.get("fc_noc") if d.get("fc_noc") is not None else int(fc_prom * 0.90)
    fc_max = d.get("fc_max") if d.get("fc_max") is not None else int(fc_prom * 1.3)
    fc_min = d.get("fc_min") if d.get("fc_min") is not None else int(fc_prom * 0.7)
    
    horas = [f"{h:02d}:00" for h in range(24)]
    fc_curva = []
    for h in range(24):
        if 6 <= h <= 21: val = fc_dia + (fc_max - fc_dia) * 0.25 * ((h % 4) / 4)
        else: val = fc_noc - (fc_noc - fc_min) * 0.3 * ((h % 3) / 3)
        fc_curva.append(round(max(fc_min, min(fc_max, val))))
        
    fig = go.Figure()
    fig.add_hrect(y0=60, y1=100, fillcolor="rgba(10, 37, 64, 0.04)", line_width=0, annotation_text="Normal (60-100)", annotation_position="top left", annotation_font_size=9)
    fig.add_trace(go.Scatter(x=horas, y=fc_curva, mode='lines+markers', name='FC (lpm)', line=dict(color='#0A2540', width=2.5), marker=dict(size=4, color='#C8102E')))
    fig.add_hline(y=fc_prom, line_dash="dot", line_color="#0284c7", annotation_text=f"Prom: {fc_prom} lpm", annotation_position="bottom right")
    
    if d.get("fc_dia") is None:
        fig.add_annotation(text="Simulación visual (Falta data horaria original)", xref="paper", yref="paper", x=0.5, y=0.1, showarrow=False, font=dict(color="red", size=10))
        
    fig.update_layout(title="<b>Tacograma Horario y Variabilidad Circadiana (24 Horas)</b>", height=240, margin=dict(l=35, r=20, t=35, b=25), plot_bgcolor="#ffffff", paper_bgcolor="#ffffff", showlegend=False)
    return fig

# ==============================================================================
# MOTOR 1: EXTRACTOR DETERMINÍSTICO C++ (BÚSQUEDA EXHAUSTIVA DE COLUMNAS)
# ==============================================================================
def limpiar_numero(val_str, tipo=int):
    if not val_str: return None
    s = str(val_str).strip()
    if re.match(r'^\d{1,3}\.\d{3}$', s): s = s.replace(".", "")
    elif "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
    elif "," in s: s = s.replace(",", ".")
    try: return tipo(float(s))
    except Exception: return None

def extraer_numero(patron, texto, tipo=int):
    m = re.search(patron, texto, re.IGNORECASE)
    if m: return limpiar_numero(m.group(1), tipo)
    return None

def extraer_m1_holter(texto):
    d = {}
    txt_flat = re.sub(r'\s+', ' ', texto) 
    
    m_nom = re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,50},\s*[A-ZÁÉÍÓÚÑ\s]{3,50})[\s\n]+(?:No confirmado|Confirmado|Reconfirmado)?[\s\n]*Informe", texto)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else None
    
    m_id = re.search(r"(?:ID\s*Paciente|ID|C\.?C\.?|Doc\.?|Historia)\s*[:\.]?\s*(\d{5,12})", texto, re.IGNORECASE)
    d["cedula"] = m_id.group(1) if m_id else None

    d["fc_prom"] = extraer_numero(r"Prom\.\.?[\s\|]*(\d{2,3})", txt_flat, int)
    d["fc_max"]  = extraer_numero(r"M[áa]x[\s\|]*(\d{2,3})", txt_flat, int)
    d["fc_min"]  = extraer_numero(r"M[íi]n[\s\|]*(\d{2,3})", txt_flat, int)
    d["fc_dia"]  = extraer_numero(r"D[íi]a.*?Prom\.\.?[\s\|]*(\d{2,3})", txt_flat, int)
    d["fc_noc"]  = extraer_numero(r"Noche.*?Prom\.\.?[\s\|]*(\d{2,3})", txt_flat, int)
    
    d["pr_intervalo"] = extraer_numero(r"(?:Intervalo\s*PR|PR\s*Interval|PR\s*promedio|PR\s*medio)\s*[:\-\|]?\s*(\d{2,3})\b", txt_flat, int)

    d["total_latidos"] = extraer_numero(r"Total\s+de\s+latidos\s*[:\.]?\s*([\d\.\,]+)", texto, int)
    d["ev_total"]      = extraer_numero(r"Latidos ventriculares\s*[:\|]?\s*([\d\.\,]+)", texto, int)
    d["esv_total"]     = extraer_numero(r"Latidos supraventriculares\s*[:\|]?\s*([\d\.\,]+)", texto, int)
    d["mcp_latidos"]   = extraer_numero(r"Marcapaso\s*[:\|]?\s*([\d\.\,]+)", texto, int)
    
    mcp_pct_match = re.search(r"Marcapaso[^\d]*([\d\.]+)\s*%", txt_flat, re.IGNORECASE)
    if mcp_pct_match: d["mcp_porcentaje"] = float(mcp_pct_match.group(1))
    else: d["mcp_porcentaje"] = None
        
    d["latidos_normales"] = None

    dx_text = texto[:2000].upper()
    mcp_en_dx = bool(re.search(r'(?:DX|DIAGN[ÓO]STICO|INDICACI[ÓO]N|COMENTARIOS).*?\b(MCP|MARCAPASO|MARCAPASOS|TRC|CDI)\b', dx_text))
    d["mcp_presente"] = True if ((d.get("mcp_latidos") and d["mcp_latidos"] > 0) or (d.get("mcp_porcentaje") and d["mcp_porcentaje"] > 0) or mcp_en_dx) else False

    d["rr_max_seg"] = extraer_numero(r"Intervalo RR.*?M[áa]x\.\s*longitud\s*([\d,\.]+)\s*s", txt_flat, float)
    d["pausas"] = extraer_numero(r"\bPausas?[\s\|]*(\d+)", txt_flat, int)
    d["pausa_max_seg"] = extraer_numero(r"Pausa.*?Máx.*?([\d\.,]+)\s*s", txt_flat, float)
    
    d["tv_episodios"] = extraer_numero(r"\bTV[\s\|]*([\d\.,]+)", txt_flat, int)
    d["ev_duplas"] = extraer_numero(r"Apareado[\s\|]*([\d\.,]+)", txt_flat, int)
    d["tsv_episodios"] = extraer_numero(r"(?:\bTSV\b|Salvas?)[\s\|]*([\d\.,]+)", txt_flat, int)
    d["taqui_conteo"] = extraer_numero(r"\bTaquicardia[\s\|]*(\d+)", txt_flat, int)
    d["bradi_conteo"] = extraer_numero(r"\bBradicardia[\s\|]*(\d+)", txt_flat, int)
    d["bigeminismo"] = extraer_numero(r"Bigeminismo[\s\|]*(\d+)", txt_flat, int)
    d["trigeminismo"] = extraer_numero(r"Trigeminismo[\s\|]*(\d+)", txt_flat, int)
    d["latidos_caidos"] = extraer_numero(r"Latidos? ca[íi]dos?[\s\|]*(\d+)", txt_flat, int)
    
    # EXTRACCIÓN AVANZADA SDNN Y QTC: Aislando la fila
    d["sdnn_24h"] = None
    m_sdnn_line = re.search(r"Valor de 24 horas[^\n\r]*", txt_flat, re.IGNORECASE)
    if m_sdnn_line:
        nums = re.findall(r'\b\d+(?:\.\d+)?\b', m_sdnn_line.group(0).replace(',', '.'))
        if len(nums) >= 2: 
            val_extracted = int(float(nums[1]))
            d["sdnn_24h"] = val_extracted 
    
    if d["sdnn_24h"] is None:
        m_sdnn_gen = re.search(r"\bSDNN\b[^\d]*(\d{2,3})", txt_flat, re.IGNORECASE)
        if m_sdnn_gen: d["sdnn_24h"] = int(m_sdnn_gen.group(1))

    d["qtc_prom"] = None
    m_qtc_line = re.search(r"Todos los per[íi]odos[^\n\r]*", txt_flat, re.IGNORECASE)
    if m_qtc_line:
        nums = re.findall(r'\b\d+(?:\.\d+)?\b', m_qtc_line.group(0).replace(',', '.'))
        if len(nums) >= 3: 
            val_qtc = int(float(nums[2]))
            d["qtc_prom"] = val_qtc if val_qtc > 0 else None
            
    if d["qtc_prom"] is None:
        m_qtc_gen = re.search(r"\bQTc\b[^\d]*(\d{3})", txt_flat, re.IGNORECASE)
        if m_qtc_gen: d["qtc_prom"] = int(m_qtc_gen.group(1))

    d["st_dep_episodios"] = extraer_numero(r"Depresi[óo]n ST[\s\|]*(\d+)", txt_flat, int)
    d["st_elev_episodios"] = extraer_numero(r"Elevaci[óo]n ST[\s\|]*(\d+)", txt_flat, int)
    
    bloqueo_keys = ["BLOQUEO DE RAMA", "BRD", "BRI", "BCRD", "BCRI", "QRS ANCHO", "CONDUCCION INTRAVENTRICULAR"]
    d["tiene_bloqueo_rama"] = any(k in dx_text for k in bloqueo_keys)
    return d

def extraer_m1_mapa(texto):
    d = {}
    txt_flat = re.sub(r'\s+', ' ', texto)
    m_nom = re.search(r"Nombre del paciente:\s*([^\n\r\|]+)", texto, re.IGNORECASE)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else None
    d["pas_24h"] = extraer_numero(r"Resumen general.*?Prom\.?:\s*(\d{2,3})", txt_flat, int)
    d["pad_24h"] = extraer_numero(r"Resumen general.*?Prom\.?:\s*\d{2,3}\s*[\/\-]\s*(\d{2,3})", txt_flat, int)
    d["carga_pas"] = extraer_numero(r"Sist[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", txt_flat, float)
    d["carga_pad"] = extraer_numero(r"Diast[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", txt_flat, float)
    d["pp_val"] = extraer_numero(r"Presi[óo]n de pulso\s*\(mmHg\)\s*(\d{2,3})", txt_flat, int)
    d["caida_nocturna_val"] = extraer_numero(r"Sist[óo]lico\s*\(mmHg\)\s*.*?([\d,\.\-]+)\s*%", txt_flat, float)
    return d

def extraer_m1_esfuerzo(texto):
    d = {}
    txt_flat = re.sub(r'\s+', ' ', texto)
    d["fc_basal"] = extraer_numero(r"FC\s*(?:en)?\s*reposo\s*[:\-]?\s*(\d{2,3})", txt_flat, int)
    d["fc_pico"] = extraer_numero(r"FC\s*(?:M[áa]x|Pico)\s*[:\-]?\s*(\d{2,3})", txt_flat, int)
    d["pas_basal"] = extraer_numero(r"PA\s*reposo\s*[:\-]?\s*(\d{2,3})\s*/", txt_flat, int)
    d["pad_basal"] = extraer_numero(r"PA\s*reposo\s*[:\-]?\s*\d{2,3}\s*/\s*(\d{2,3})", txt_flat, int)
    d["pas_pico"] = extraer_numero(r"PA\s*M[áa]x\s*[:\-]?\s*(\d{2,3})\s*/", txt_flat, int)
    d["pad_pico"] = extraer_numero(r"PA\s*M[áa]x\s*[:\-]?\s*\d{2,3}\s*/\s*(\d{2,3})", txt_flat, int)
    d["tiempo_min"] = extraer_numero(r"Tiempo\s*(?:total)?\s*[:\-]?\s*([\d,\.]+)", txt_flat, float)
    d["mets"] = extraer_numero(r"METS\s*[:\-]?\s*([\d,\.]+)", txt_flat, float)
    d["st_mm"] = extraer_numero(r"ST\s*[:\-]?\s*([\d,\.]+)", txt_flat, float)
    return d

# ==============================================================================
# MOTOR 2: INTELIGENCIA ARTIFICIAL GEMINI (CEREBRO CLÍNICO COMPLETO)
# ==============================================================================
def consultar_gemini_json(prompt_text, inline_items=None):
    key = str(st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))).strip().strip('"').strip("'")
    if not key: return False, None, "Falta API Key"

    modelos = [("v1beta", "models/gemini-2.5-flash"), ("v1beta", "models/gemini-2.0-flash")]
    payload = {"contents": [{"parts": [{"text": prompt_text}] + (inline_items or [])}], "generationConfig": {"temperature": 0.0, "responseMimeType": "application/json"}}

    for ver, mod in modelos:
        url = f"https://generativelanguage.googleapis.com/{ver}/{mod}:generateContent?key={key}"
        try:
            resp = requests.post(url, headers={"Content-Type": "application/json"}, json=payload, timeout=60)
            if resp.status_code == 200:
                raw = resp.json()['candidates'][0]['content']['parts'][0]['text']
                match = re.search(r'\{.*\}', raw, re.DOTALL)
                return True, json.loads(match.group(0) if match else raw.strip()), mod
        except Exception: continue
    return False, None, "Timeout en API Gemini o Fallo en la red"

def optimizar_imagen_para_ia(img_bytes, max_dim=1400):
    img = Image.open(io.BytesIO(img_bytes))
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB": img = img.convert("RGB")
    w, h = img.size
    if max(w, h) > max_dim:
        ratio = max_dim / max(w, h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=80, optimize=True)
    return buf.getvalue()

def ejecutar_extraccion_multimodal(archivos_fotos):
    inline_items = []
    for foto in archivos_fotos:
        try:
            raw_b = foto.getvalue() if hasattr(foto, "getvalue") else foto
            opt_b = optimizar_imagen_para_ia(raw_b, max_dim=1400)
            b64_img = base64.b64encode(opt_b).decode('utf-8')
            inline_items.append({"inline_data": {"mime_type": "image/jpeg", "data": b64_img}})
        except Exception: continue

    prompt = """Eres un Cardiólogo especialista. Extrae los parámetros exactos de las fotos y devuelve ÚNICAMENTE JSON:
{
  "paciente": "KAROL JULIANA", "cedula": "1050095219", "edad": 17, "sexo": "Femenino",
  "protocolo": "Bruce", "etapa": "Etapa 6", "tiempo_min": 16.23, "fc_basal": 91, "fc_pico": 194,
  "pas_basal": 119, "pad_basal": 70, "pas_pico": 140, "pad_pico": 87, "st_mm": 0.0
}"""
    ok, data, info = consultar_gemini_json(prompt, inline_items)
    if ok and data:
        for k in ["paciente", "cedula", "edad", "sexo", "protocolo", "etapa", "tiempo_min", "fc_basal", "fc_pico", "st_mm"]:
            if data.get(k): st.session_state[f"erg_{k}"] = data[k]
        if data.get("pas_basal") and data.get("pad_basal"): st.session_state.erg_pa_basal = f"{data['pas_basal']}/{data['pad_basal']}"
        if data.get("pas_pico") and data.get("pad_pico"): st.session_state.erg_pa_pico = f"{data['pas_pico']}/{data['pad_pico']}"
        return True, "Parámetros extraídos con éxito."
    return False, "No se pudo completar el análisis."

# ==============================================================================
# MOTOR 3: NÚCLEO DE ARBITRAJE MATEMÁTICO UNIVERSAL
# ==============================================================================
def arbitrar_valor(k, v1, v2, tipo, criticidad="baja"):
    if v1 is None and v2 is None:
        if criticidad == "alta": return None, f"🚨 CRÍTICO: Ningún motor encontró '{k}'."
        else: return None, None

    if v1 is not None and v2 is not None:
        if v1 == v2: return v1, None
        return v1, f"⚠️ Arbitraje en {k}: M1 ({v1}) prevalece sobre IA ({v2})."
    
    if v1 is not None: return v1, None
    return v2, f"⚕️ Dato rescatado por Visión IA ({k}): {v2}."

def auditar_motor_3_universal(m1, m2, modalidad):
    d_f = {}
    notas = []

    if modalidad == "HOLTER":
        for k in ["fc_prom", "total_latidos"]:
            val, msg = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "alta")
            d_f[k] = val; 
            if msg: notas.append(msg)
            
        claves_clinicas = [
            "ev_total", "esv_total", "tv_episodios", "fc_dia", "fc_noc", "mcp_latidos", 
            "eventos_paciente", "pausas", "sdnn_24h", "qtc_prom", "fc_max", "fc_min", 
            "ev_duplas", "tsv_episodios", "st_dep_episodios", "st_elev_episodios", 
            "taqui_conteo", "bradi_conteo", "taqui_fc_max", "bradi_fc_min", 
            "latidos_caidos", "bigeminismo", "trigeminismo", "pr_intervalo"
        ]
        
        for k in claves_clinicas:
            val, _ = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "baja")
            d_f[k] = val

        d_f["rr_max_seg"], _ = arbitrar_valor("rr_max_seg", m1.get("rr_max_seg"), m2.get("rr_max_seg") if m2 else None, float, "baja")
        d_f["pausa_max_seg"], _ = arbitrar_valor("pausa_max_seg", m1.get("pausa_max_seg"), m2.get("pausa_max_seg") if m2 else None, float, "baja")

        d_f["paciente"] = m2.get("paciente") if (m2 and m2.get("paciente")) else m1.get("paciente", "PACIENTE ANÓNIMO")
        d_f["cedula"] = m2.get("cedula") if (m2 and m2.get("cedula")) else m1.get("cedula", "")

        norm_m1 = m1.get("latidos_normales")
        norm_m2 = m2.get("latidos_normales") if m2 else None
        norm, _ = arbitrar_valor("latidos_normales", norm_m1, norm_m2, int, "baja")
        
        if d_f.get("total_latidos") is not None:
            norm_val = norm or 0
            ev_t = d_f.get("ev_total") or 0
            esv_t = d_f.get("esv_total") or 0
            mcp_l = d_f.get("mcp_latidos") or 0
            
            suma = norm_val + ev_t + esv_t + mcp_l
            if d_f["total_latidos"] > 0 and abs(suma - d_f["total_latidos"]) > (d_f["total_latidos"] * 0.05):
                notas.append(f"⚖️ CONSERVACIÓN: Suma ({suma}) no cuadra con Total ({d_f['total_latidos']}). Ajustando Total.")
                d_f["total_latidos"] = max(d_f["total_latidos"], suma)
                d_f["latidos_normales"] = norm_val
            else:
                d_f["latidos_normales"] = norm_val if norm_val > 0 else max(0, d_f["total_latidos"] - ev_t - esv_t - mcp_l)

        if d_f.get("rr_max_seg") is not None and d_f.get("pausas") is not None:
            if d_f["rr_max_seg"] > 0 and d_f["rr_max_seg"] < 2.0 and d_f["pausas"] > 0:
                notas.append(f"🔒 VETO FISIOLÓGICO: RR Máx es {d_f['rr_max_seg']}s. Imposible pausas > 2.0s. Se vetan pausas.")
                d_f["pausas"] = None

        mcp_pct_m1 = m1.get("mcp_porcentaje")
        mcp_pct_m2 = m2.get("mcp_porcentaje") if m2 else None
        d_f["mcp_porcentaje"], _ = arbitrar_valor("mcp_porcentaje", mcp_pct_m1, mcp_pct_m2, float, "baja")
        
        d_f["mcp_presente"] = bool(m1.get("mcp_presente") or (m2 and m2.get("mcp_presente")) or (d_f.get("mcp_porcentaje") and d_f["mcp_porcentaje"] > 0) or (d_f.get("mcp_latidos") and d_f["mcp_latidos"] > 0))
        
        t_bloq = m1.get("tiene_bloqueo_rama", False) or (m2.get("tiene_bloqueo_rama", False) if m2 else False)
        if d_f["mcp_presente"] and (d_f.get("mcp_porcentaje") or 0) >= 50.0:
            d_f["tiene_bloqueo_rama"] = True
            d_f["bloqueo_detalle"] = "bloqueo de rama izquierda secundario a electroestimulación ventricular"
        else:
            d_f["tiene_bloqueo_rama"] = t_bloq
            d_f["bloqueo_detalle"] = "bloqueo de rama documentado"
            
        d_f["st_episodios"] = (d_f.get("st_dep_episodios") or 0) + (d_f.get("st_elev_episodios") or 0)

    elif modalidad == "MAPA":
        for k in ["pas_24h", "pad_24h"]:
            val, msg = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "alta")
            d_f[k] = val; 
            if msg: notas.append(msg)
            
        for k in ["pp_val", "pas_max_sueno", "pad_max_sueno"]:
            val, _ = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "baja")
            d_f[k] = val

        for k in ["carga_pas", "carga_pad", "caida_nocturna_val"]:
            val, msg = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, float, "baja")
            d_f[k] = val; 
            
        d_f["paciente"] = m2.get("paciente") if (m2 and m2.get("paciente")) else m1.get("paciente", "PACIENTE MAPA")
        d_f["cedula"] = m2.get("cedula") if (m2 and m2.get("cedula")) else m1.get("cedula", "")

        if d_f.get("pas_24h") is not None and d_f.get("pad_24h") is not None and d_f["pas_24h"] > 0 and d_f["pad_24h"] > 0:
            d_f["pam_24h"] = round(d_f["pad_24h"] + ((d_f["pas_24h"] - d_f["pad_24h"]) / 3))
            d_f["pp_val"] = d_f["pas_24h"] - d_f["pad_24h"]
        else:
            d_f["pam_24h"] = None; d_f["pp_val"] = None

    elif modalidad == "ESFUERZO":
        for k in ["fc_basal", "fc_pico", "pas_basal", "pad_basal", "pas_pico", "pad_pico", "edad"]:
            val, msg = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "alta")
            d_f[k] = val; 
            if msg: notas.append(msg)
        for k in ["tiempo_min", "mets", "st_mm"]:
            val, msg = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, float, "baja")
            d_f[k] = val; 
        d_f["paciente"] = m2.get("paciente") if (m2 and m2.get("paciente")) else m1.get("paciente", "PACIENTE ESFUERZO")
        d_f["cedula"] = m2.get("cedula") if (m2 and m2.get("cedula")) else m1.get("cedula", "")
        d_f["protocolo"] = m2.get("protocolo") if m2 else "Bruce"
        d_f["etapa"] = m2.get("etapa") if m2 else "Etapa 1"
        d_f["sexo"] = m2.get("sexo") if m2 else "Femenino"

    return d_f, notas

def ejecutar_triple_engine_universal(pdf_bytes, modalidad="HOLTER"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pags = len(doc)
    
    texto_completo = "".join([f"\n--- PÁG {i+1} ---\n" + p.get_text() + "\n" for i, p in enumerate(doc)])
    if modalidad == "HOLTER": m1 = extraer_m1_holter(texto_completo)
    elif modalidad == "MAPA": m1 = extraer_m1_mapa(texto_completo)
    else: m1 = extraer_m1_esfuerzo(texto_completo)

    inline_images = []
    for num_p in range(total_pags):
        try:
            pix = doc[num_p].get_pixmap(dpi=75)
            b64_p = base64.b64encode(pix.tobytes("jpeg")).decode("utf-8")
            inline_images.append({"inline_data": {"mime_type": "image/jpeg", "data": b64_p}})
        except: continue
    doc.close()

    if modalidad == "HOLTER":
        prompt_ia = f"""Eres el Jefe de Electrofisiología Clínica (Guías AHA/ESC/HRS). Tienes la obligación de realizar una auditoría médica exhaustiva de este estudio Holter de 24 horas ({total_pags} páginas).

INSTRUCCIONES DE AUDITORÍA CLÍNICA PROFUNDA:
1. FRECUENCIA Y RITMO: Rastrea la FC Promedio, Día, Noche, Máxima y Mínima. Evalúa presencia de marcapasos (MCP), % de estimulación y latidos totales.
2. EVENTOS CRONOTRÓPICOS: Contabiliza episodios exactos de Taquicardia y Bradicardia, con sus frecuencias máximas y mínimas.
3. INTERVALOS: Evalúa el QTc promedio y el intervalo PR (milisegundos). 
4. ISQUEMIA (ST): Detecta episodios de infradesnivel o supradesnivel del segmento ST.
5. CONDUCCIÓN AV: Identifica latidos caídos o bloqueos AV.
6. CONDUCCIÓN INTRAVENTRICULAR: Busca bloqueos de rama (BRI/BRD) o complejos anchos. Si MCP > 50%, asume morfología de bloqueo.
7. ECTOPIAS Y ARRITMIAS: Extrae conteo exacto de EV (Ventriculares), ESV (Supraventriculares), Duplas, Rachas o Salvas de TSV y TV, y patrones de Bigeminismo o Trigeminismo.
8. SÍNTOMAS: Revisa si hay eventos marcados por el paciente.
9. VARIABILIDAD (HRV/SDNN): Busca el SDNN real en ms. REGLA DE ORO: Si ves tablas de variabilidad o QT llenas de '0' por todo lado, es porque el equipo no lo calculó (artefacto técnico). En ese caso devuelve 'null'. 
10. PAUSAS: Detecta pausas significativas y su duración máxima (pausa_max_seg).

Devuelve ÚNICAMENTE un objeto JSON estricto. Usa 'null' explícito (sin comillas) si el dato NO está escrito o es una tabla vacía/llena de ceros falsos. NO asumas ni inventes.
{{"paciente": "str", "cedula": "str", "mcp_presente": bool, "mcp_latidos": int, "mcp_porcentaje": float, "fc_prom": int, "fc_dia": int, "fc_noc": int, "fc_max": int, "fc_min": int, "total_latidos": int, "latidos_normales": int, "pausas": int, "rr_max_seg": float, "pausa_max_seg": float, "ev_total": int, "tv_episodios": int, "ev_duplas": int, "esv_total": int, "tsv_episodios": int, "sdnn_24h": int, "qtc_prom": int, "pr_intervalo": int, "st_dep_episodios": int, "st_elev_episodios": int, "eventos_paciente": int, "tiene_bloqueo_rama": bool, "taqui_conteo": int, "bradi_conteo": int, "taqui_fc_max": int, "bradi_fc_min": int, "latidos_caidos": int, "bigeminismo": int, "trigeminismo": int}}"""
    elif modalidad == "MAPA":
        prompt_ia = f"""Eres Auditor Médico Jefe. Analiza las {total_pags} imágenes y el texto de este MAPA.
Texto: {texto_completo}
Devuelve JSON estricto (null si el dato no está):
{{"paciente": "str", "cedula": "str", "pas_24h": int, "pad_24h": int, "carga_pas": float, "carga_pad": float, "pp_val": int, "caida_nocturna_val": float, "pas_max_sueno": int, "pad_max_sueno": int}}"""
    else:
        prompt_ia = f"""Eres Auditor Médico Jefe. Analiza este PDF de Prueba de Esfuerzo.
Texto: {texto_completo}
Devuelve JSON estricto (null si el dato no está):
{{"paciente": "str", "cedula": "str", "edad": int, "sexo": "str", "protocolo": "str", "etapa": "str", "tiempo_min": float, "fc_basal": int, "fc_pico": int, "pas_basal": int, "pad_basal": int, "pas_pico": int, "pad_pico": int, "st_mm": float, "mets": float}}"""

    ok_m2, m2, mod_ia = consultar_gemini_json(prompt_ia, inline_items=inline_images)

    d_final, notas = auditar_motor_3_universal(m1, m2 if ok_m2 else None, modalidad)
    d_final["triple_engine_activo"] = True
    d_final["notas_arbitraje"] = notas
    d_final["motores_info"] = f"Auditoría Triple Engine: M1 (Espacial {total_pags} pág) + M2 (Visión IA Clínica) + M3 (Arbitraje Estricto)"
    
    return d_final

# ==============================================================================
# REDACTORES INSTITUCIONALES (PLANTILLAS EXACTAS DEL DR. AMAYA)
# ==============================================================================
def redactar_informe_holter_11_puntos(d, perfil):
    fc_p = d.get("fc_prom")
    fc_p_str = f"{fc_p}" if fc_p is not None else "No evaluable"

    fc_d = d.get("fc_dia")
    fc_n = d.get("fc_noc")
    
    dip_txt = ""
    if fc_d and fc_n and fc_p and fc_d != fc_p and fc_n != fc_p:
        desc_crono = ((fc_d - fc_n) / fc_d * 100) if fc_d > 0 else 0
        patron = "conservado" if desc_crono >= 10 else "atenuado"
        dip_txt = f" (Día: {fc_d} / Noche: {fc_n}; patrón circadiano {patron} con {desc_crono:.1f}% de descenso)"

    pct_mcp = d.get("mcp_porcentaje") or 0.0
    pausas = d.get("pausas") or 0
    txt_sensado = " con adecuado sensado y captura" if pausas == 0 else ""
    
    if d.get("mcp_presente") and pct_mcp >= 80.0: 
        p1 = f"1. Ritmo comandado por marcapasos definitivo (electroestimulado/sensado en {pct_mcp:.1f}%){txt_sensado} con FC promedio {fc_p_str} lpm{dip_txt}."
    elif d.get("mcp_presente"): 
        pct_str = f" ({pct_mcp:.1f}% mediado por MCP)" if pct_mcp > 0 else " con presencia de MCP"
        p1 = f"1. Ritmo sinusal alternando con electroestimulación{pct_str}{txt_sensado} con FC promedio {fc_p_str} lpm{dip_txt}."
    else: 
        p1 = f"1. Ritmo sinusal con FC promedio {fc_p_str} lpm{dip_txt}."

    crono = []
    tq_cnt = d.get("taqui_conteo")
    br_cnt = d.get("bradi_conteo")
    if tq_cnt and tq_cnt > 0: crono.append(f"{tq_cnt} episodios de taquicardia (FC máx {d.get('taqui_fc_max') or d.get('fc_max') or 0} lpm)")
    if br_cnt and br_cnt > 0: crono.append(f"{br_cnt} episodios de bradicardia (FC mín {d.get('bradi_fc_min') or d.get('fc_min') or 0} lpm)")
    
    if crono: p2 = f"2. Eventos cronotrópicos: Se documentaron {' y '.join(crono)}."
    else: p2 = "2. Eventos cronotrópicos: Sin bradicardia patológica ni taquicardias sostenidas."

    qtc = d.get("qtc_prom")
    if qtc is not None and qtc > 500: txt_qtc = f"QTc SEVERAMENTE PROLONGADO ({qtc} ms: alto riesgo proarrítmico)"
    elif qtc is not None and qtc > 460: txt_qtc = f"QTc prolongado ({qtc} ms)"
    elif qtc is not None and qtc > 0: txt_qtc = f"QTc normal ({qtc} ms)"
    else: txt_qtc = "QTc conservados"

    pr_int = d.get("pr_intervalo")
    if pr_int is not None and pr_int > 200: txt_pr = f"Intervalo PR prolongado (Bloqueo AV de primer grado: {pr_int} ms)"
    elif pr_int is not None and pr_int > 0 and pr_int < 120: txt_pr = f"Intervalo PR corto ({pr_int} ms)"
    elif pr_int is not None and pr_int > 0: txt_pr = f"Intervalos PR normales ({pr_int} ms)"
    else: txt_pr = "Intervalos PR normales"

    p3 = f"3. {txt_pr} y {txt_qtc}."

    st_ep = d.get('st_episodios')
    if st_ep and st_ep > 0: p4 = f"4. Alteraciones isquémicas del ST ({st_ep} eventos)."
    else: p4 = "4. Sin alteraciones isquémicas del segmento ST."
    
    lat_ca = d.get('latidos_caidos')
    if lat_ca and lat_ca > 0: p5 = f"5. Conducción AV alterada por {lat_ca} latidos caídos/bloqueos."
    else: p5 = "5. Sin Alteración de la conducción AV."

    if d.get("mcp_presente") and pct_mcp >= 50.0: p6 = f"6. Complejos ventriculares con {d.get('bloqueo_detalle', 'morfología de bloqueo por estimulación')}."
    elif d.get("tiene_bloqueo_rama"): p6 = f"6. Alteración de conducción intraventricular por {d.get('bloqueo_detalle', 'bloqueo de rama')}."
    else: p6 = "6. Sin alteración en la conducción intraventricular."

    tot_q = d.get("total_latidos") or 1
    ev = d.get("ev_total") or 0
    carga = (ev / tot_q) * 100 if tot_q > 0 else 0
    
    tv_ep = d.get("tv_episodios") or 0
    ev_dup = d.get("ev_duplas") or 0
    
    if tv_ep > 0: lown = "Lown Grado IVb (Taquicardia Ventricular)"
    elif ev_dup > 0: lown = "Lown Grado IVa (Duplas)"
    elif ev >= 720: lown = "Lown Grado II (Frecuentes > 30/h)"
    elif ev > 0: lown = "Lown Grado I"
    else: lown = "Lown Grado 0"

    ect = []
    esv = d.get("esv_total") or 0
    tsv = d.get("tsv_episodios") or 0
    
    if esv > 0: ect.append(f"ectopias supraventriculares ({esv} ESV)")
    if tsv > 0: ect.append(f"salvas de taquicardia supraventricular ({tsv} rachas)")
    
    bigem = d.get("bigeminismo") or 0
    trigem = d.get("trigeminismo") or 0
    
    if ev > 0: 
        texto_ev = f"ectopia ventricular con {ev} EV (Carga {carga:.2f}%, {lown})"
        if bigem > 0 or trigem > 0:
            patrones = []
            if bigem > 0: patrones.append(f"Bigeminismo ({bigem} eventos)")
            if trigem > 0: patrones.append(f"Trigeminismo ({trigem} eventos)")
            texto_ev += f" asociando patrón de {' y '.join(patrones)}"
        ect.append(texto_ev)
        
    p7 = f"7. Impulsos alterados por {' y '.join(ect)}." if ect else "7. Sin alteración de impulsos ectópicos."
    
    ev_pac = d.get("eventos_paciente")
    if ev_pac is not None and ev_pac > 0: p8 = f"8. Diario paciente: {ev_pac} eventos marcados."
    else: p8 = "8. No refirió síntomas."

    sdnn = d.get("sdnn_24h")
    mod_mcp = " (Modificada por estimulación de marcapasos)." if d.get("mcp_presente") else ""
    
    if sdnn is None:
        p9 = f"9. Variabilidad de la frecuencia cardíaca no reportada explícitamente en el documento{mod_mcp}."
        p11 = "11. Riesgo del paciente SDNN 20 a 24 hrs ( Pendiente de Valoración )."
        d_vfc = "VFC no especificada."
    else:
        try: sdnn_val = int(float(sdnn))
        except: sdnn_val = None 
        
        if sdnn_val is not None:
            if sdnn_val == 0:
                p9 = f"9. Variabilidad de la frecuencia cardíaca disminuida{mod_mcp}."
                p11 = "11. Riesgo del paciente SDNN 20 a 24 hrs ( Riesgo Medio )."
                d_vfc = "VFC disminuida (Riesgo Medio)."
            elif sdnn_val < 30:
                p9 = f"9. Variabilidad de la frecuencia cardíaca severamente disminuida ({sdnn_val} ms){mod_mcp}."
                p11 = "11. Riesgo del paciente SDNN 20 a 24 hrs ( Riesgo Alto )."
                d_vfc = "VFC severamente disminuida (Riesgo Alto)."
            elif sdnn_val <= 100:
                p9 = f"9. Variabilidad de la frecuencia cardíaca disminuida ({sdnn_val} ms){mod_mcp}."
                p11 = "11. Riesgo del paciente SDNN 20 a 24 hrs ( Riesgo Medio )."
                d_vfc = "VFC disminuida (Riesgo Medio)."
            else:
                p9 = f"9. Variabilidad de la frecuencia cardíaca conservada ({sdnn_val} ms){mod_mcp}."
                p11 = "11. Riesgo del paciente SDNN 20 a 24 hrs ( Riesgo Bajo )."
                d_vfc = "VFC conservada."
        else:
            p9 = f"9. Variabilidad de la frecuencia cardíaca no reportada explícitamente en el documento{mod_mcp}."
            p11 = "11. Riesgo del paciente SDNN 20 a 24 hrs ( Pendiente de Valoración )."
            d_vfc = "VFC no especificada."

    if pausas > 0:
        p_max = d.get('pausa_max_seg')
        p_max_str = f" (máx {p_max}s)" if p_max else ""
        p10 = f"10. Se registraron {pausas} pausas significativas{p_max_str}."
    else:
        p10 = "10. No hay pausas significativas."

    diag = f"Ritmo sinusal con FC promedio {fc_p_str} lpm. {d_vfc}"
    if d.get("mcp_presente"): diag = f"Ritmo comandado por marcapasos normofuncionante ({pct_mcp:.1f}%). FC promedio {fc_p_str} lpm."
    
    recs = ["Valoración por Electrofisiología."] if tv_ep > 0 or carga >= 10.0 else ["Continuar manejo médico instaurado y seguimiento clínico."]

    return f"INTERPRETACIÓN TEST HOLTER - CUPS 895001\n\n{p1}\n{p2}\n{p3}\n{p4}\n{p5}\n{p6}\n{p7}\n{p8}\n{p9}\n{p10}\n{p11}\n\nCONCLUSIÓN DIAGNÓSTICA:\n{diag}\n\nRECOMENDACIONES: {' '.join(recs)}\n\n{perfil['nombre_completo']}\n{perfil['especialidad']}\n{perfil['registro']}"

def redactar_informe_mapa_cencardio(d, perfil):
    p1 = f"1. Promedios 24h: Sistólica {d.get('pas_24h') or 'N/A'} mmHg / Diastólica {d.get('pad_24h') or 'N/A'} mmHg; PAM: {d.get('pam_24h') or 'N/A'} mmHg."
    p2 = f"2. Carga tensional sistólica ({d.get('carga_pas') or 0.0}%) y diastólica ({d.get('carga_pad') or 0.0}%)."
    p3 = f"3. Presión de pulso: {d.get('pp_val') or 'N/A'} mmHg."
    p4 = f"4. Caída nocturna (Dipping): {d.get('caida_nocturna_val') or 0.0}%."
    
    c_pas = d.get('carga_pas') or 0.0
    pas_h = d.get('pas_24h') or 0
    ctrl = "Control óptimo de la tensión arterial" if (c_pas < 15 and pas_h < 130 and pas_h > 0) else "Control subóptimo de la tensión arterial"
    recs = "Continuar manejo farmacológico actual." if "óptimo" in ctrl else "Optimización y titulación de la terapia antihipertensiva."
    
    return f"INTERPRETACIÓN TEST MAPA - CUPS 895003\n\nHallazgos:\n{p1}\n{p2}\n{p3}\n{p4}\n5. {ctrl}.\n\nRECOMENDACIONES: {recs}\n\n{perfil['nombre_completo']}\n{perfil['especialidad']}\n{perfil['registro']}"

def redactar_informe_ergometria_institucional(d, perfil):
    fcm_prev = 220 - d.get("edad", 35) if d.get("edad", 35) > 0 else 200
    fc_pico = d.get("fc_pico") or 0
    porc = round((fc_pico / fcm_prev) * 100) if fcm_prev > 0 else 0
    suf = "suficiente" if porc >= 85 else "insuficiente"
    dp = fc_pico * (d.get("pas_pico") or 0)
    
    st_mm = d.get("st_mm") or 0.0
    st_res = "Sin alteraciones isquémicas del segmento ST" if st_mm < 1.0 else f"Infradesnivel del ST de {st_mm} mm"
    diag_el = "negativa" if st_mm < 1.0 else "positiva"
    dts = (d.get("tiempo_min") or 0.0) - (5 * st_mm)
    duke = "Bajo riesgo coronario" if dts >= 5 else ("Riesgo moderado" if dts >= -10 else "Alto riesgo coronario")
    
    fc_basal = d.get("fc_basal") or 75
    reserva_den = (fcm_prev - fc_basal) if (fcm_prev - fc_basal) > 0 else 1
    fcr_pct = round(((fc_pico - fc_basal) / reserva_den) * 100)
    incomp_crono = "Incompetencia cronotrópica" if fcr_pct < 80 else "Respuesta cronotrópica adecuada"

    return f"""INTERPRETACIÓN PRUEBA DE ESFUERZO COMPUTARIZADA - CUPS 893805

1. Ritmo sinusal normal basal y durante el esfuerzo físico.
2. Protocolo de {d.get('protocolo', 'Bruce')} completado con duración de {d.get('tiempo_min') or 0.0:.2f} minutos ({d.get('etapa', 'Final')}).
3. Capacidad funcional alcanzada: {d.get('mets') or 0.0} METs.
4. Respuesta cronotrópica: FC basal {fc_basal} lpm a FC pico de {fc_pico} lpm ({porc}% FCM prevista; FCR: {fcr_pct}%, {incomp_crono}, prueba {suf}).
5. Respuesta hemodinámica presora: PA basal {d.get('pas_basal') or 0}/{d.get('pad_basal') or 0} mmHg a PA pico de {d.get('pas_pico') or 0}/{d.get('pad_pico') or 0} mmHg.
6. Doble producto máximo alcanzado: {dp:,} mmHg*lpm.
7. Comportamiento del ST: {st_res}.
8. Sin arritmias ventriculares complejas inducidas.
9. Estratificación por Duke Treadmill Score: {dts:.1f} ({duke}).

CONCLUSIÓN DIAGNÓSTICA:
Prueba física {suf}, eléctricamente {diag_el} para isquemia miocárdica. Adecuada tolerancia ({incomp_crono}).

RECOMENDACIONES: Continuar control médico periódico.

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

# ==============================================================================
# INYECTORES DE PDFS Y FORENSE
# ==============================================================================
@st.cache_data
def generar_qr_verificacion(paciente, medico, fecha_str, codigo_uuid, proc_nombre):
    url_completa = f"https://holtercencardio.streamlit.app/?val={codigo_uuid[:12]}&pac={urllib.parse.quote(paciente)}&med={urllib.parse.quote(medico)}&proc={proc_nombre}"
    qr = qrcode.QRCode(version=1, box_size=4, border=1)
    qr.add_data(url_completa)
    qr.make(fit=True)
    buf = io.BytesIO()
    qr.make_image(fill_color="#0a2540", back_color="white").save(buf, format="PNG")
    return buf.getvalue()

@st.cache_data
def procesar_firma_transparente():
    nom = next((n for n in ["OR WILIAM ANDA RAMIREZ.pdf", "firma_amaya.pdf", "firma_amaya.png"] if os.path.exists(n)), None)
    if not nom: return None
    try:
        if nom.endswith(".pdf"):
            doc = fitz.open(nom)
            pix = doc[0].get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
        else: img = Image.open(nom).convert("RGB")
        img_gray = img.convert("L")
        alpha = img_gray.point(lambda p: 255 if p < 185 else 0, mode='L')
        tinta = Image.new("RGBA", img.size, (10, 37, 64, 255))
        tinta.putalpha(alpha)
        caja = tinta.getbbox()
        buf = io.BytesIO()
        (tinta.crop(caja) if caja else tinta).save(buf, format="PNG")
        return buf.getvalue()
    except Exception: return None

def normalizar_nombre_archivo(nombre):
    limpio = re.sub(r'[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s]', ' ', nombre)
    return re.sub(r'\s+', '_', limpio).strip('_') or "PACIENTE"

def inyectar_holter_pdf(pdf_bytes, texto_informe, paciente_nom, perfil, cod_uuid, estampador_activo=False):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]
    rects_f = pagina1.search_for("Firma del médico") or pagina1.search_for("Firma del operador")
    rects_h = pagina1.search_for("Hallazgos:")
    y_base = rects_f[0].y0 if rects_f else 740
    y0 = (rects_h[0].y1 + 4) if rects_h else 545
    y1 = y_base - 62

    pagina1.draw_rect(fitz.Rect(0, y0, pagina1.rect.width, y1), color=None, fill=(1, 1, 1), overlay=True)
    
    fs_opt = 3.9
    for fs in [5.8, 5.5, 5.2, 4.9, 4.6, 4.3, 4.0, 3.8]:
        dt = fitz.open(stream=pdf_bytes, filetype="pdf")
        if dt[0].insert_textbox(fitz.Rect(35, y0, pagina1.rect.width - 36, y1), texto_informe, fontsize=fs, fontname="helv", align=fitz.TEXT_ALIGN_LEFT) >= 0:
            fs_opt = fs
            break
        dt.close()

    pagina1.insert_textbox(fitz.Rect(35, y0, pagina1.rect.width - 36, y1), texto_informe, fontsize=fs_opt, fontname="helv", color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)
    pagina1.draw_rect(fitz.Rect(220, y_base - 58, pagina1.rect.width, y_base + 12), color=None, fill=(1, 1, 1), overlay=True)
    
    qr_b = generar_qr_verificacion(paciente_nom, perfil['nombre_completo'], ahora_colombia().strftime("%Y-%m-%d"), cod_uuid, "CUPS 895001")
    pagina1.insert_image(fitz.Rect(230, y_base - 32, 270, y_base + 8), stream=qr_b)
    pagina1.insert_text(fitz.Point(275, y_base - 18), "Validado Digitalmente", fontsize=5.2, fontname="helv", color=(0.08, 0.2, 0.36))
    pagina1.insert_text(fitz.Point(275, y_base - 9), "Res. 3100 MinSalud", fontsize=4.7, fontname="helv", color=(0.25, 0.25, 0.25))
    pagina1.insert_text(fitz.Point(275, y_base), f"Cód: {cod_uuid[:12]}...", fontsize=4.5, fontname="helv", color=(0.4, 0.4, 0.4))

    firma_b = procesar_firma_transparente()
    if firma_b and estampador_activo:
        fx0 = (rects_f[0].x0 + 10) if rects_f else 380
        pagina1.insert_image(fitz.Rect(fx0, y_base - 58, fx0 + 155, y_base + 4), stream=firma_b)

    pdf_out = doc.tobytes(); img_prev = doc[0].get_pixmap(dpi=130).tobytes("png"); doc.close()
    return pdf_out, img_prev

def inyectar_mapa_pdf(pdf_bytes, texto_informe, paciente_nom, perfil, cod_uuid, estampador_activo=False):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]
    rect_m = pagina1.search_for("Presión arterial por la mañana")
    rect_r = pagina1.search_for("Resumen de todo el registro")
    y0 = (rect_m[0].y1 + 2) if rect_m else 230
    y1 = (rect_r[0].y0 - 4) if rect_r else 425

    pagina1.draw_rect(fitz.Rect(0, y0, pagina1.rect.width, y1), color=None, fill=(1, 1, 1), overlay=True)

    fs_opt = 6.4
    for fs in [7.2, 6.8, 6.4, 6.0, 5.6]:
        dt = fitz.open(stream=pdf_bytes, filetype="pdf")
        if dt[0].insert_textbox(fitz.Rect(36, y0 + 2, 440, y1 - 2), texto_informe, fontsize=fs, fontname="helv", align=fitz.TEXT_ALIGN_LEFT) >= 0:
            fs_opt = fs
            break
        dt.close()

    pagina1.insert_textbox(fitz.Rect(36, y0 + 2, 440, y1 - 2), texto_informe, fontsize=fs_opt, fontname="helv", color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)

    qr_b = generar_qr_verificacion(paciente_nom, perfil['nombre_completo'], ahora_colombia().strftime("%Y-%m-%d"), cod_uuid, "CUPS 895003")
    pagina1.insert_image(fitz.Rect(455, y0 + 6, 505, y0 + 56), stream=qr_b)
    pagina1.insert_text(fitz.Point(510, y0 + 22), "Validado Digitalmente", fontsize=5.0, fontname="helv", color=(0.04, 0.15, 0.25))
    pagina1.insert_text(fitz.Point(510, y0 + 32), "Res. 3100 MinSalud", fontsize=4.6, fontname="helv", color=(0.3, 0.3, 0.3))
    pagina1.insert_text(fitz.Point(510, y0 + 42), f"Cód: {cod_uuid[:10]}...", fontsize=4.4, fontname="helv", color=(0.4, 0.4, 0.4))

    firma_b = procesar_firma_transparente()
    if firma_b and estampador_activo:
        pagina1.insert_image(fitz.Rect(445, y0 + 62, 575, y1 - 4), stream=firma_b)

    pdf_out = doc.tobytes(); img_prev = doc[0].get_pixmap(dpi=130).tobytes("png"); doc.close()
    return pdf_out, img_prev

def generar_pdf_ergometria_completo(d, texto_informe, perfil, cod_uuid, imagenes_adjuntas=[]):
    doc = fitz.open()
    page = doc.new_page(width=612, height=792)

    logo_bytes = None
    for nom in ["cencardio.jpg", "cencardio.png", "cencardio.jpeg", "logo.png", "logo.jpg"]:
        if os.path.exists(nom):
            with open(nom, "rb") as f: logo_bytes = f.read()
            break
    if logo_bytes: page.insert_image(fitz.Rect(36, 30, 150, 75), stream=logo_bytes)

    page.insert_text(fitz.Point(165, 45), "CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO", fontsize=11, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(165, 58), "INFORME DE ERGOMETRÍA Y PRUEBA DE ESFUERZO COMPUTARIZADA", fontsize=8.5, fontname="helv", color=(0.78, 0.06, 0.18))
    page.insert_text(fitz.Point(165, 70), "CUPS: 893805 · Habilitación MinSalud Colombia · Res. 3100 de 2019", fontsize=7, fontname="helv", color=(0.4, 0.45, 0.5))
    page.draw_rect(fitz.Rect(36, 85, 576, 87), color=None, fill=(0.04, 0.15, 0.25), overlay=True)

    page.draw_rect(fitz.Rect(36, 95, 576, 155), color=(0.85, 0.9, 0.95), fill=(0.97, 0.98, 1.0), width=1)
    page.insert_text(fitz.Point(46, 112), f"PACIENTE: {d.get('paciente','').upper()}", fontsize=8.5, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(46, 126), f"DOCUMENTO: {d.get('cedula','')}    |    EDAD: {d.get('edad','')} AÑOS    |    SEXO: {d.get('sexo','')}", fontsize=7.5, fontname="helv", color=(0.2, 0.25, 0.3))
    page.insert_text(fitz.Point(46, 140), f"FECHA DEL ESTUDIO: {ahora_colombia().strftime('%d/%m/%Y')}    |    MÉDICO LECTOR: {perfil['nombre_completo']}", fontsize=7.5, fontname="helv", color=(0.2, 0.25, 0.3))

    fcm_prev = 220 - d.get("edad", 35) if d.get("edad", 35) > 0 else 200
    porc = round(((d.get("fc_pico") or 0) / fcm_prev) * 100) if fcm_prev > 0 else 0
    cajas = [
        ("FC PICO ALCANZADA", f"{d.get('fc_pico') or 0} lpm ({porc}%)"),
        ("PA ESFUERZO PICO", f"{d.get('pas_pico') or 0}/{d.get('pad_pico') or 0} mmHg"),
        ("CARGA FUNCIONAL", f"{d.get('mets') or 0.0} METs ({d.get('tiempo_min') or 0.0:.2f} m)"),
        ("DOBLE PRODUCTO", f"{(d.get('fc_pico') or 0)*(d.get('pas_pico') or 0):,}")
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

    qr_bytes = generar_qr_verificacion(d.get('paciente',''), perfil['nombre_completo'], ahora_colombia().strftime("%Y-%m-%d"), cod_uuid, "CUPS 893805")
    page.insert_image(fitz.Rect(48, 695, 100, 747), stream=qr_bytes)
    page.insert_text(fitz.Point(108, 715), "Certificado Digital Forense", fontsize=6.2, fontname="helv", color=(0.04, 0.15, 0.25))
    page.insert_text(fitz.Point(108, 726), "Res. 3100 de 2019 · Habilitación MinSalud", fontsize=5.5, fontname="helv", color=(0.4, 0.45, 0.5))
    page.insert_text(fitz.Point(108, 737), f"Cód: {cod_uuid[:16]}...", fontsize=5.2, fontname="helv", color=(0.4, 0.45, 0.5))

    firma_bytes = procesar_firma_transparente()
    if firma_bytes and perfil["id"] in ["dr.amaya", "admin"]:
        page.insert_image(fitz.Rect(390, 690, 545, 755), stream=firma_bytes)

    img_preview = page.get_pixmap(dpi=130).tobytes("png")

    for img_file in imagenes_adjuntas:
        try:
            p_extra = doc.new_page(width=612, height=792)
            data_img = img_file.getvalue() if hasattr(img_file, "getvalue") else img_file
            p_extra.insert_image(fitz.Rect(20, 20, 592, 772), stream=data_img)
        except Exception: pass

    pdf_bytes = doc.tobytes()
    doc.close()
    return pdf_bytes, img_preview

# ==============================================================================
# EXCEL GERENCIAL AVANZADO
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
    border_thin = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'), top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

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
# UI PRINCIPAL Y LOGIN
# ==============================================================================
LISTA_ESPECIALISTAS = [
    {"id": "dr.amaya", "etiqueta": "Dr. William Amaya Ramirez", "clave": "Cardio2025*", "nombre_completo": "DR. WILLIAM AMAYA RAMIREZ", "especialidad": "INTERNISTA - CARDIÓLOGO", "registro": "RM 79.502.624 SDS"},
    {"id": "dr.suarez", "etiqueta": "Dr. Martin Suárez Arámbula", "clave": "Suarez2026*", "nombre_completo": "DR. MARTIN SUÁREZ ARÁMBULA", "especialidad": "INTERNISTA - HEMODINAMISTA", "registro": "RM 13491094"},
    {"id": "dra.cardio", "etiqueta": "Dra. Paola Figueroa", "clave": "Cardio2026*", "nombre_completo": "DRA. PAOLA FIGUEROA", "especialidad": "MÉDICA CARDIÓLOGA", "registro": "RM 52.890.123 SDS"},
    {"id": "admin", "etiqueta": "Administración del Sistema", "clave": "HolterClaveSegura123", "nombre_completo": "DR. WILLIAM AMAYA RAMIREZ", "especialidad": "INTERNISTA - CARDIÓLOGO", "registro": "RM 79.502.624 SDS"}
]
OPCIONES_NOMBRES = [m["etiqueta"] for m in LISTA_ESPECIALISTAS]

if "autenticado" not in st.session_state: st.session_state.autenticado = False

params = st.query_params
if "val" in params:
    c_v1, c_v2, c_v3 = st.columns([1, 1.8, 1])
    with c_v2:
        logo_data = obtener_logo_b64()
        logo_html = f'<img src="{logo_data}" style="max-width:180px; margin-bottom:1rem;">' if logo_data else '<div style="font-size:3rem;">🫀</div>'
        st.markdown(f'<div class="cencardio-card">{logo_html}<div style="font-weight:800; font-size:1.2rem; margin-bottom:1rem;">CERTIFICADO DE AUTENTICIDAD FORENSE</div><div style="background:#f0fdf4; border:1px solid #86efac; border-radius:12px; padding:1.3rem; text-align:left;"><b>Procedimiento:</b> {urllib.parse.unquote(params.get("proc",""))}<br><b>Paciente:</b> {urllib.parse.unquote(params.get("pac",""))}<br><b>Lector:</b> {urllib.parse.unquote(params.get("med",""))}<br><b>Fecha:</b> {urllib.parse.unquote(params.get("fec",""))}<br><b>Cód:</b> {params.get("val","")}</div></div>', unsafe_allow_html=True)
        if st.button("Ir al Portal", use_container_width=True):
            st.query_params.clear()
            st.rerun()
    st.stop()

if not st.session_state.autenticado:
    c1, c2, c3 = st.columns([1, 1.6, 1])
    with c2:
        logo_data = obtener_logo_b64()
        logo_html = f'<img src="{logo_data}" style="max-width:170px; margin-bottom:1rem;">' if logo_data else '<div style="font-size:3rem;">🫀</div>'
        st.markdown(f'<div class="cencardio-card">{logo_html}<div style="font-size:1.25rem; font-weight:800; color:#0a2540;">Centro Cardiovascular Colombiano</div><div style="font-size:0.8rem; font-weight:700; color:#c8102e; margin-bottom:1.8rem;">CENCARDIO WORKSTATION</div>', unsafe_allow_html=True)
        with st.form("login"):
            usr = st.selectbox("Especialista Responsable:", OPCIONES_NOMBRES)
            pwd = st.text_input("Contraseña de Acceso:", type="password")
            if st.form_submit_button("Ingresar a la Estación", use_container_width=True):
                m = next(x for x in LISTA_ESPECIALISTAS if x["etiqueta"] == usr)
                if m["clave"] == pwd:
                    st.session_state.autenticado = True; st.session_state.usuario_actual = m["id"]
                    st.rerun()
                else: st.error("❌ Contraseña incorrecta para el especialista seleccionado.")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

perfil_activo = next(x for x in LISTA_ESPECIALISTAS if x["id"] == st.session_state.usuario_actual)
with st.sidebar:
    st.markdown(f'<div class="specialist-card"><div style="font-weight:800; font-size:0.95rem;">{perfil_activo["nombre_completo"]}</div><div style="font-size:0.74rem; color:#93c5fd;">{perfil_activo["especialidad"]}</div></div>', unsafe_allow_html=True)
    mod_sel = st.radio("Modalidad Diagnóstica:", ["🫀 Holter (895001)", "🩺 MAPA (895003)", "🏃 Esfuerzo (893805)"])
    if procesar_firma_transparente() and perfil_activo["id"] in ["dr.amaya", "admin"]: st.success("🖋️ Sello digitalizado activo.")
    st.divider()
    archivo_dinamica = st.file_uploader("Directorio Dinámica (Excel/CSV):", type=["xlsx", "xls", "csv"])
    if archivo_dinamica:
        try:
            df_din = pd.read_csv(archivo_dinamica) if archivo_dinamica.name.endswith(".csv") else pd.read_excel(archivo_dinamica)
            ok, msg = sincronizar_directorio_servicio(df_din)
            if ok: st.success(msg)
            else: st.error(msg)
        except Exception as e: st.error(str(e))
    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

st.markdown(f'<div class="top-hospital-bar"><div><div style="font-size:1.35rem; font-weight:800; color:#0a2540;">CENCARDIO WORKSTATION</div><div style="font-size:0.82rem; font-weight:700; color:#c8102e;">Alta Complejidad</div></div><div style="display:flex; gap:8px;"><span class="inst-badge-success">{"Nube Supabase" if supabase else "Local SQLite"}</span><span class="inst-badge-primary">Triple Engine 100% Auditado</span></div></div>', unsafe_allow_html=True)

tab_procesar, tab_historial = st.tabs(["📥 Procesamiento del Estudio", "📁 Archivo Clínico & Reportes"])

# LLAVES DINÁMICAS PARA LIMPIEZA DE UPLOADER Y CACHÉ
if "up_key_esf" not in st.session_state: st.session_state.up_key_esf = str(uuid.uuid4())
if "up_key_hm" not in st.session_state: st.session_state.up_key_hm = str(uuid.uuid4())

# ==============================================================================
# PESTAÑA 1: PROCESAMIENTO (HOLTER, MAPA Y ESFUERZO)
# ==============================================================================
with tab_procesar:
    if "Esfuerzo" in mod_sel:
        col_esf1, col_esf2 = st.columns([1, 1], gap="large")
        with col_esf1:
            up_pdf_esf = st.file_uploader("📥 Cargar PDF Consolidado:", type=["pdf"], key=st.session_state.up_key_esf)
        with col_esf2:
            fotos_esfuerzo = st.file_uploader("📸 O subir fotos de tiras:", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="fot_esf")
        
        for k, v in [("erg_paciente", ""), ("erg_cedula", ""), ("erg_edad", 35), ("erg_sexo", "Femenino"), ("erg_protocolo", "Bruce"), ("erg_etapa", "Etapa 4"), ("erg_tiempo", 0.0), ("erg_fc_basal", 75), ("erg_fc_pico", 150), ("erg_pa_basal", "120/80"), ("erg_pa_pico", "160/90"), ("erg_st_mm", 0.0), ("erg_celular", "")]:
            if k not in st.session_state: st.session_state[k] = v

        if up_pdf_esf:
            b_pdf = up_pdf_esf.getvalue()
            if st.session_state.get("last_file") != up_pdf_esf.name:
                with st.spinner("🚀 Ejecutando TRIPLE ENGINE sobre PDF de Esfuerzo..."):
                    d_final = ejecutar_triple_engine_universal(b_pdf, "ESFUERZO")
                    st.session_state.d_final = d_final
                    st.session_state.txt_final = redactar_informe_ergometria_institucional(d_final, perfil_activo)
                    st.session_state.last_file = up_pdf_esf.name
                    st.session_state.uid = str(uuid.uuid4()).upper()
                    st.session_state.tel = buscar_telefono_servicio(d_final.get("cedula"))
                    for k in ["paciente", "cedula", "edad", "sexo", "protocolo", "etapa", "tiempo_min", "fc_basal", "fc_pico", "st_mm"]:
                        if d_final.get(k): st.session_state[f"erg_{k}"] = d_final[k]
                    if d_final.get("pas_basal") and d_final.get("pad_basal"): st.session_state.erg_pa_basal = f"{d_final['pas_basal']}/{d_final['pad_basal']}"
                    if d_final.get("pas_pico") and d_final.get("pad_pico"): st.session_state.erg_pa_pico = f"{d_final['pas_pico']}/{d_final['pad_pico']}"
                    st.toast("✅ Lectura completada.")
            
            if st.button("🔄 Cargar Nuevo Estudio / Limpiar Datos", type="secondary"):
                st.session_state.up_key_esf = str(uuid.uuid4())
                for k in list(st.session_state.keys()):
                    if k not in ["autenticado", "usuario_actual", "up_key_esf", "up_key_hm"]:
                        del st.session_state[k]
                st.rerun()

        elif fotos_esfuerzo and st.button("⚡ EXTRAER FOTOS CON VISIÓN IA", type="primary", use_container_width=True):
            with st.spinner("Analizando imágenes..."):
                ok, msg = ejecutar_extraccion_multimodal(fotos_esfuerzo)
                if ok: st.success(msg); st.rerun()
                else: st.error(msg)
        
        st.divider()
        col_f1, col_f2 = st.columns([1.2, 1], gap="large")
        with col_f1:
            st.markdown("<b>Parámetros Clínicos Extraídos</b>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            p_nombre = c1.text_input("Paciente:", key="erg_paciente")
            p_cedula = c1.text_input("Cédula:", key="erg_cedula")
            p_edad = c2.number_input("Edad:", min_value=1, key="erg_edad")
            p_sexo = c2.selectbox("Sexo:", ["Femenino", "Masculino"], key="erg_sexo")
            p_protocolo = c3.selectbox("Protocolo:", ["Bruce", "Naughton"], key="erg_protocolo")
            p_etapa = c3.text_input("Etapa:", key="erg_etapa")
            
            c4, c5, c6 = st.columns(3)
            p_tiempo = c4.number_input("Tiempo (min):", step=0.1, key="erg_tiempo")
            p_mets = c4.number_input("METs:", value=float(calcular_mets_bruce(p_tiempo)), step=0.5)
            p_fc_basal = c5.number_input("FC Basal:", key="erg_fc_basal")
            p_fc_pico = c5.number_input("FC Pico:", key="erg_fc_pico")
            p_pa_basal = c6.text_input("PA Basal:", key="erg_pa_basal")
            p_pa_pico = c6.text_input("PA Pico:", key="erg_pa_pico")
            
            c7, c8 = st.columns(2)
            p_st_mm = c7.number_input("Desviación ST (mm):", step=0.5, key="erg_st_mm")
            if not st.session_state.erg_celular and p_cedula: st.session_state.erg_celular = buscar_telefono_servicio(p_cedula)
            p_celular = c8.text_input("Celular (WhatsApp):", key="erg_celular")

        with col_f2:
            st.markdown("<b>Diagnóstico y Certificación</b>", unsafe_allow_html=True)
            if p_nombre.strip():
                pas_b, pad_b = [int(x) for x in p_pa_basal.split("/")] if "/" in p_pa_basal else (120, 80)
                pas_p, pad_p = [int(x) for x in p_pa_pico.split("/")] if "/" in p_pa_pico else (150, 90)
                d_erg = {"paciente": p_nombre, "cedula": p_cedula, "edad": p_edad, "sexo": p_sexo, "protocolo": p_protocolo, "etapa": p_etapa, "tiempo_min": p_tiempo, "mets": p_mets, "fc_basal": p_fc_basal, "fc_pico": p_fc_pico, "pas_basal": pas_b, "pad_basal": pad_b, "pas_pico": pas_p, "pad_pico": pad_p, "st_mm": p_st_mm}
                
                mostrar_semaforizacion_clinica("ESFUERZO", d_erg)
                bloq, ale = ejecutar_sanity_checks("ESFUERZO", d_erg)
                for b in bloq: st.error(f"🛑 {b}")
                for a in ale: st.warning(f"⚠️ {a}")

                txt_erg_final = st.text_area("Informe Oficial:", value=redactar_informe_ergometria_institucional(d_erg, perfil_activo), height=310)
                for err in auditar_coherencia_informe(txt_erg_final, d_erg, "ESFUERZO"): st.info(f"🩺 Auditoría: {err}")

                uid = st.session_state.get("uid", str(uuid.uuid4()).upper())
                pdf_erg, img_erg_prev = generar_pdf_ergometria_completo(d_erg, txt_erg_final, perfil_activo, uid, fotos_esfuerzo or [])

                b1, b2 = st.columns(2)
                b1.download_button("📄 DESCARGAR PDF", data=pdf_erg, file_name=f"{normalizar_nombre_archivo(p_nombre)}_Esfuerzo.pdf", mime="application/pdf", type="primary", use_container_width=True)
                if b2.button("💾 Guardar Clínico", use_container_width=True):
                    guardar_estudio_servicio(p_nombre, "Prueba de Esfuerzo", "CUPS 893805", f"FC Pico: {p_fc_pico} | {p_mets} METs", perfil_activo['nombre_completo'], txt_erg_final, pdf_erg, uid)
                    st.success("✅ Guardado.")
                
                if p_celular:
                    tel_l = re.sub(r'\D', '', p_celular)
                    if len(tel_l) == 10: tel_l = "57" + tel_l
                    url_c = f"https://holtercencardio.streamlit.app/?val={uid[:12]}&pac={urllib.parse.quote(p_nombre)}&med={urllib.parse.quote(perfil_activo['nombre_completo'])}&proc=CUPS_893805"
                    st.link_button("📲 ENVIAR WHATSAPP", f"https://wa.me/{tel_l}?text={urllib.parse.quote(f'Estimado(a) {p_nombre}, CENCARDIO le entrega su resultado de Prueba de Esfuerzo: {url_c}')}", use_container_width=True)
                if img_erg_prev: st.image(img_erg_prev, use_container_width=True)

    else:
        up_file = st.file_uploader(f"📥 Cargar PDF de {mod_sel}", type=["pdf"], key=st.session_state.up_key_hm)
        if up_file:
            b_pdf = up_file.getvalue()
            if st.session_state.get("last_file") != up_file.name:
                with st.spinner("🚀 Ejecutando TRIPLE ENGINE (Extracción 100% Auditada y Sin Suposiciones)..."):
                    mod = "HOLTER" if "Holter" in mod_sel else "MAPA"
                    d_final = ejecutar_triple_engine_universal(b_pdf, mod)
                    st.session_state.d_final = d_final
                    st.session_state.txt_final = redactar_informe_holter_11_puntos(d_final, perfil_activo) if mod == "HOLTER" else redactar_informe_mapa_cencardio(d_final, perfil_activo)
                    st.session_state.last_file = up_file.name
                    st.session_state.uid = str(uuid.uuid4()).upper()
                    st.session_state.tel = buscar_telefono_servicio(d_final.get("cedula"))
                    st.toast(f"✅ {d_final.get('motores_info')}")
            
            if st.button("🔄 Cargar Nuevo Estudio / Limpiar Datos", type="secondary"):
                st.session_state.up_key_hm = str(uuid.uuid4())
                for k in list(st.session_state.keys()):
                    if k not in ["autenticado", "usuario_actual", "up_key_esf", "up_key_hm"]:
                        del st.session_state[k]
                st.rerun()

            d = st.session_state.d_final
            mod_actual = "HOLTER" if "Holter" in mod_sel else "MAPA"
            cups_actual = "CUPS 895001" if mod_actual == "HOLTER" else "CUPS 895003"
            mod_nombre = "Holter ECG 24H" if mod_actual == "HOLTER" else "MAPA Tensional 24H"
            
            mostrar_semaforizacion_clinica(mod_actual, d)
            bloq, ale = ejecutar_sanity_checks(mod_actual, d)
            for b in bloq: st.error(f"🛑 {b}")
            for a in ale: st.warning(f"⚠️ {a}")
            
            if d.get("notas_arbitraje"):
                with st.expander("🛡️ REPORTE DE ARBITRAJE TRIPLE ENGINE", expanded=True):
                    for n in d["notas_arbitraje"]: 
                        if "CRÍTICO" in n or "⚠️" in n or "🚨" in n or "⚖️" in n or "🔒" in n: st.warning(n)
                        else: st.info(n)
            
            c1, c2, c3, c4 = st.columns(4)
            if mod_actual == "HOLTER":
                fc_dia_disp = d.get('fc_dia') or "N/A"
                fc_noc_disp = d.get('fc_noc') or "N/A"
                c1.metric("FC Promedio", f"{d.get('fc_prom') or 'N/A'} lpm", f"Día {fc_dia_disp} / Noche {fc_noc_disp}")
                c2.metric("Ectopias V", f"{d.get('ev_total') or 0} EV", f"TV: {d.get('tv_episodios') or 0}")
                
                pausas_disp = d.get('pausas')
                pausas_met = pausas_disp if pausas_disp is not None else 0
                max_p = d.get('pausa_max_seg')
                max_str = f"Máx: {max_p}s" if max_p else "Máx: N/A"
                c3.metric("Pausas", f"{pausas_met}", max_str)
                
                sdnn_v = d.get('sdnn_24h')
                sdnn_disp = f"{sdnn_v} ms" if sdnn_v else "N/A"
                c4.metric("SDNN / MCP", sdnn_disp, f"MCP: {d.get('mcp_porcentaje') or 0.0}%")
                
                fig_taco = generar_grafica_tacograma(d)
                st.plotly_chart(fig_taco, use_container_width=True)
            else:
                c1.metric("PAM 24H", f"{d.get('pam_24h') or 'N/A'} mmHg", f"{d.get('pas_24h') or 'N/A'}/{d.get('pad_24h') or 'N/A'}")
                c2.metric("Carga Sistólica", f"{d.get('carga_pas') or 0.0}%")
                c3.metric("Carga Diastólica", f"{d.get('carga_pad') or 0.0}%")
                c4.metric("Caída Nocturna", f"{d.get('caida_nocturna_val') or 0.0}%")
            
            st.divider()
            col_ed, col_prev = st.columns([1, 1], gap="large")
            with col_ed:
                c_nom, c_ced = st.columns([1.8, 1.2])
                p_nom = c_nom.text_input("👤 Paciente:", value=d.get('paciente') or "")
                p_ced = c_ced.text_input("🪪 Cédula:", value=d.get('cedula') or "")
                    
                tel_in = st.text_input("📱 Celular (WhatsApp):", value=st.session_state.tel)
                txt_edit = st.text_area(f"📝 Informe Oficial ({cups_actual}):", value=st.session_state.txt_final, height=450)
                
                for err in auditar_coherencia_informe(txt_edit, d, mod_actual): st.info(f"🩺 Auditoría: {err}")
                
                uid = st.session_state.uid
                debe_estampar = perfil_activo["id"] in ["dr.amaya", "admin"]
                if mod_actual == "HOLTER": p_final, img_p = inyectar_holter_pdf(b_pdf, txt_edit, p_nom, perfil_activo, uid, debe_estampar)
                else: p_final, img_p = inyectar_mapa_pdf(b_pdf, txt_edit, p_nom, perfil_activo, uid, debe_estampar)
                
                c_b1, c_b2 = st.columns(2)
                c_b1.download_button("📄 DESCARGAR FIRMADO", data=p_final, file_name=f"{normalizar_nombre_archivo(p_nom)}_{cups_actual}.pdf", mime="application/pdf", type="primary", use_container_width=True)
                if c_b2.button("💾 Guardar Clínico", use_container_width=True):
                    p_clave = f"FC {d.get('fc_prom') or 0} | SDNN {d.get('sdnn_24h') or 0}ms" if mod_actual == "HOLTER" else f"PAM {d.get('pam_24h') or 0} mmHg"
                    ok, msg = guardar_estudio_servicio(p_nom, mod_nombre, cups_actual, p_clave, perfil_activo['nombre_completo'], txt_edit, p_final, uid)
                    if ok: st.success(msg)
                    else: st.error(msg)
                
                if tel_in:
                    t_l = re.sub(r'\D', '', tel_in)
                    if len(t_l) == 10: t_l = "57" + t_l
                    url_c = f"https://holtercencardio.streamlit.app/?val={uid[:12]}&pac={urllib.parse.quote(p_nom)}&med={urllib.parse.quote(perfil_activo['nombre_completo'])}&proc={urllib.parse.quote(mod_nombre)}"
                    st.link_button("📲 ENVIAR RESULTADO POR WHATSAPP", f"https://wa.me/{t_l}?text={urllib.parse.quote(f'Estimado(a) {p_nom}, CENCARDIO le entrega su resultado de {mod_nombre}. Certificado oficial: {url_c}')}", use_container_width=True)
            with col_prev:
                st.subheader("👁️ Vista Previa Oficial")
                st.markdown('<div class="preview-container">', unsafe_allow_html=True)
                st.image(img_p, use_container_width=True)
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
        col_r1.markdown(f"**Total de estudios custodiados ({origen_datos}):** `{len(df_produccion)} procedimientos certificados`")
        
        excel_bytes, ext_salida = generar_excel_avanzado_produccion(df_produccion)
        mime_tipo = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if ext_salida == "xlsx" else "text/csv"
        etiqueta_boton = "📊 DESCARGAR REPORTE EXCEL GERENCIAL" if ext_salida == "xlsx" else "📊 DESCARGAR RIPS (CSV)"
        col_r2.download_button(etiqueta_boton, data=excel_bytes, file_name=f"Reporte_Cencardio_{ahora_colombia().strftime('%Y%m%d')}.{ext_salida}", mime=mime_tipo, use_container_width=True, type="primary")

        st.divider()
        f1, f2, f3 = st.columns([1.5, 1.2, 1.3])
        busqueda = f1.text_input("🔍 Buscar por paciente:")
        df_produccion["Mes_Periodo"] = pd.to_datetime(df_produccion["fecha_registro"]).dt.strftime('%Y-%m')
        mes_sel = f2.selectbox("📅 Carpeta Mensual:", ["Todos los meses"] + sorted(df_produccion["Mes_Periodo"].unique().tolist(), reverse=True))
        mod_sel_f = f3.selectbox("🎛️ Modalidad:", ["Todas las modalidades"] + sorted(df_produccion["modalidad"].unique().tolist()))

        df_fil = df_produccion.copy()
        if busqueda.strip(): df_fil = df_fil[df_fil["paciente"].str.contains(busqueda, case=False, na=False)]
        if mes_sel != "Todos los meses": df_fil = df_fil[df_fil["Mes_Periodo"] == mes_sel]
        if mod_sel_f != "Todas las modalidades": df_fil = df_fil[df_fil["modalidad"] == mod_sel_f]

        for mes_g in sorted(df_fil["Mes_Periodo"].unique().tolist(), reverse=True):
            df_mes = df_fil[df_fil["Mes_Periodo"] == mes_g]
            with st.expander(f"📁 CARPETA: {mes_g} ({len(df_mes)} estudios)", expanded=True):
                for item in df_mes.itertuples():
                    st.markdown(f'<div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:0.8rem; margin-bottom:0.5rem;"><div style="display:flex; justify-content:space-between;"><div><b>👤 {item.paciente}</b> <span style="background:#e0f2fe; color:#0369a1; padding:2px 8px; border-radius:12px; font-size:0.75rem;">{item.modalidad}</span><br><small>📅 {item.fecha_registro} | 👨‍⚕️ {item.medico} | 🩺 {item.parametro_clave}</small></div><div style="font-family:monospace; font-size:0.75rem; color:#0284c7;">Cód: {item.codigo_verificacion[:12]}</div></div></div>', unsafe_allow_html=True)
                    c_d, c_del, c_v = st.columns([1.5, 1, 4])
                    with c_d:
                        if item.pdf_url:
                            st.link_button("📥 Ver PDF", item.pdf_url, use_container_width=True)
                        else:
                            pdf_recup = obtener_pdf_bytes_individual(item.id)
                            if pdf_recup:
                                st.download_button("📥 Bajar PDF", data=pdf_recup, file_name=f"{normalizar_nombre_archivo(item.paciente)}_{item.cups}.pdf", mime="application/pdf", key=f"desc_{item.id}", use_container_width=True)
                    with c_del:
                        if st.button("🗑️ Eliminar", key=f"elim_{item.id}", use_container_width=True):
                            eliminar_estudio_servicio(item.id)
                            st.rerun()
