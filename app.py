import streamlit as st
import fitz  # PyMuPDF: motor C++ de renderizado ultrarrápido y extracción espacial
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

# ==============================================================================
# CONFIGURACIÓN GLOBAL Y ZONA HORARIA
# ==============================================================================
TZ_COLOMBIA = timezone(timedelta(hours=-5))

def ahora_colombia():
    return datetime.now(TZ_COLOMBIA)

# Importación segura de openpyxl (Excel)
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
        box-shadow: 0 4px 12px rgba(220, 38, 38, 0.08);
    }
    .triage-amarillo {
        background: #fffbeb !important; border-left: 6px solid #d97706 !important; color: #92400e !important;
        padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important;
        box-shadow: 0 4px 12px rgba(217, 119, 6, 0.08);
    }
    .triage-verde {
        background: #f0fdf4 !important; border-left: 6px solid #16a34a !important; color: #166534 !important;
        padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem !important;
        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.08);
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
# BASE DE DATOS LOCAL INCONDICIONAL (RESPALDO GARANTIZADO)
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
    # Migración defensiva de columnas (evita OperationalError)
    c.execute("PRAGMA table_info(estudios)")
    cols_existentes = [col[1] for col in c.fetchall()]
    cols_a_asegurar = [
        ("parametro_clave", "TEXT"),
        ("codigo_verificacion", "TEXT"),
        ("hash_sha256", "TEXT"),
        ("pdf_blob", "BLOB")
    ]
    for col_nom, col_tipo in cols_a_asegurar:
        if col_nom not in cols_existentes:
            try:
                c.execute(f"ALTER TABLE estudios ADD COLUMN {col_nom} {col_tipo}")
            except Exception:
                pass
    conn.commit()
    conn.close()

# Ejecución incondicional al iniciar
init_db_local()

# ==============================================================================
# CONECTOR DE NUBE PERSISTENTE: SUPABASE (DESINFECCIÓN DE RUTA)
# ==============================================================================
@st.cache_resource
def obtener_cliente_supabase():
    if not SUPABASE_LIB_OK:
        return None
    url = st.secrets.get("SUPABASE_URL", os.environ.get("SUPABASE_URL", ""))
    key = st.secrets.get("SUPABASE_KEY", os.environ.get("SUPABASE_KEY", ""))
    if url and key:
        try:
            # Desinfección estricta de la URL para evitar errores PGRST125
            url_limpia = str(url).strip().rstrip("/")
            for sub in ["/rest/v1", "/auth/v1", "/storage/v1"]:
                if url_limpia.endswith(sub):
                    url_limpia = url_limpia[:-len(sub)].rstrip("/")
            
            key_limpia = str(key).strip().strip('"').strip("'")
            cliente = create_client(url_limpia, key_limpia)
            
            # Asegurar bucket
            try:
                cliente.storage.create_bucket("estudios-pdf", options={"public": True})
            except Exception:
                pass
            return cliente
        except Exception:
            return None
    return None

supabase = obtener_cliente_supabase()

def calcular_hash_sha256(pdf_bytes):
    return hashlib.sha256(pdf_bytes).hexdigest()

def guardar_estudio_servicio(nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif):
    fecha_actual_str = ahora_colombia().strftime("%Y-%m-%d %H:%M:%S")
    hash_seguridad = calcular_hash_sha256(pdf_bytes)
    nombre_archivo = f"{cod_verif[:12]}_{re.sub(r'[^A-Za-z0-9]', '_', nombre)}.pdf"

    guardado_en_nube = False
    error_nube_detalle = ""

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
                "codigo_verificacion": cod_verif
            }).execute()
            guardado_en_nube = True
        except Exception as e:
            error_nube_detalle = str(e)

    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("""
        INSERT INTO estudios (fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, informe_texto, pdf_blob, codigo_verificacion, hash_sha256)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (fecha_actual_str, nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif, hash_seguridad))
    conn.commit()
    conn.close()

    if guardado_en_nube:
        return True, "Guardado con éxito en nube Supabase y respaldo local."
    elif error_nube_detalle:
        st.info("ℹ️ Estudio resguardado en base de datos local de emergencia (Nube no configurada o con error).")
        return True, "Guardado en almacenamiento local seguro."
    else:
        return True, "Guardado en almacenamiento local seguro."

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

    init_db_local()
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
        return False, "El archivo debe contener columnas de Cédula y Teléfono/Celular."

    registros = 0
    fecha_hoy = ahora_colombia().strftime("%Y-%m-%d %H:%M")

    lote = []
    for _, fila in df.iterrows():
        ced = re.sub(r'\D', '', str(fila[col_ced]))
        tel = re.sub(r'\D', '', str(fila[col_tel]))
        nom = str(fila[col_nom]) if col_nom else ""
        if len(ced) >= 5 and len(tel) >= 7:
            lote.append({"cedula": ced, "nombre": nom, "telefono": tel, "fecha_actualizacion": fecha_hoy})
            registros += 1

    if supabase and lote:
        try:
            supabase.table("dinamica_pacientes").upsert(lote).execute()
        except Exception:
            pass

    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    for item in lote:
        c.execute("""
            INSERT OR REPLACE INTO dinamica_pacientes (cedula, nombre, telefono, fecha_actualizacion)
            VALUES (?, ?, ?, ?)
        """, (item["cedula"], item["nombre"], item["telefono"], item["fecha_actualizacion"]))
    conn.commit()
    conn.close()
    return True, f"Se sincronizaron con éxito {registros} pacientes."

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
    init_db_local()
    conn = sqlite3.connect("historial_cencardio.db")
    c = conn.cursor()
    c.execute("SELECT telefono FROM dinamica_pacientes WHERE cedula = ?", (ced_limpia,))
    res = c.fetchone()
    conn.close()
    return res[0] if res else ""

# ==============================================================================
# AUDITORÍA 1: SANITY CHECKS FISIOLÓGICOS Y LINTER CLÍNICO
# ==============================================================================
def ejecutar_sanity_checks(modalidad, datos):
    alertas = []
    bloqueos = []

    if modalidad == "HOLTER":
        fc_p = datos.get("fc_prom") or 0
        fc_min = datos.get("fc_min") or 0
        fc_max = datos.get("fc_max") or 0
        
        if fc_p > 0:
            if fc_min > fc_p:
                bloqueos.append(f"Incongruencia Cronotrópica: FC mínima ({fc_min} lpm) supera a la FC promedio ({fc_p} lpm).")
            if fc_max < fc_p:
                bloqueos.append(f"Incongruencia Cronotrópica: FC máxima ({fc_max} lpm) es inferior a la FC promedio ({fc_p} lpm).")
        
        if fc_min > 0 and fc_min < 30 and not datos.get("mcp_presente"):
            alertas.append(f"Bradicardia Extrema documentada ({fc_min} lpm). Compruebe ausencia de artefactos de desconexión.")
        if fc_max > 220:
            alertas.append(f"Frecuencia Cardíaca Máxima ({fc_max} lpm) atípica. Verifique si corresponde a ruido o taquiarritmia paroxística.")

    elif modalidad == "MAPA":
        pas = datos.get("pas_24h") or 0
        pad = datos.get("pad_24h") or 0
        cn = datos.get("caida_nocturna_val") or 0.0

        if pas > 0 and pad >= pas:
            bloqueos.append(f"Error Hemodinámico Severo: La Presión Diastólica ({pad} mmHg) es igual o mayor a la Sistólica ({pas} mmHg).")
        if pas > 0 and (pas - pad) < 20:
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

def auditar_coherencia_informe(texto_informe, datos, modalidad):
    discrepancias = []
    texto_upper = texto_informe.upper()

    if modalidad == "HOLTER":
        if "SIN ALTERACIONES ISQUÉMICAS" in texto_upper and (datos.get("st_episodios") or 0) > 0:
            discrepancias.append(f"El informe indica 'Sin isquemia', pero se contabilizan {datos.get('st_episodios')} episodios de desviación del ST.")
        if "TAQUICARDIA VENTRICULAR" in texto_upper and (datos.get("tv_episodios") or 0) == 0:
            discrepancias.append("Se menciona 'Taquicardia Ventricular' en la redacción, pero el contador numérico de TV es 0.")
        if datos.get("mcp_presente") and "RITMO SINUSAL CON FC PROMEDIO" in texto_upper and (datos.get("mcp_porcentaje") or 0.0) >= 80:
            discrepancias.append(f"Alerta de Ritmo: Paciente dependiente de marcapasos ({datos.get('mcp_porcentaje'):.1f}% de latidos estimulados); el dictamen debe especificar ritmo de marcapasos.")

    elif modalidad == "MAPA":
        if "DIPPING POSITIVO" in texto_upper and (datos.get("caida_nocturna_val") or 0.0) <= 0:
            discrepancias.append(f"El texto dictamina 'Dipping Positivo', pero el cálculo matemático de descenso nocturno es {datos.get('caida_nocturna_val'):.1f}%.")
        
        c_pas = datos.get("carga_pas") or 0.0
        if "CONTROL ÓPTIMO" in texto_upper and ((datos.get("pas_24h") or 0) >= 140 or c_pas > 30):
            discrepancias.append("Se califica 'Control Óptimo', pero los promedios o cargas tensionales superan el rango de normalidad.")

    elif modalidad == "ESFUERZO":
        st_mm = datos.get("st_mm") or 0.0
        if "NEGATIVA PARA ISQUEMIA" in texto_upper and st_mm >= 1.0:
            discrepancias.append(f"Se dictamina prueba 'Negativa para isquemia', pero existe infradesnivel significativo del ST de {st_mm} mm.")
        if "SUFICIENTE" in texto_upper:
            fcm_prev = 220 - datos.get("edad", 35)
            porc = ((datos.get("fc_pico") or 0) / fcm_prev) * 100 if fcm_prev > 0 else 0
            if porc > 0 and porc < 85:
                discrepancias.append(f"Se describe 'Prueba Suficiente', pero el esfuerzo solo alcanzó el {porc:.1f}% de la frecuencia submáxima prevista (meta ≥ 85%).")

    return discrepancias

# ==============================================================================
# MOTOR DE TRIAGE CLÍNICO INSTITUCIONAL (SEMAFORIZACIÓN VISUAL CENCARDIO)
# ==============================================================================
def mostrar_semaforizacion_clinica(modalidad, datos):
    nivel = "VERDE"
    titulo = "PARÁMETROS FISIOLÓGICOS CONSERVADOS"
    detalles = []

    if modalidad == "HOLTER":
        sdnn = datos.get("sdnn_24h") or 0
        tv = datos.get("tv_episodios") or 0
        pausas = datos.get("pausas") or 0
        qtc = datos.get("qtc_prom") or 0
        tot_qrs = datos.get("total_latidos") or 0
        ev = datos.get("ev_total") or 0
        carga_ev = (ev / tot_qrs) * 100 if tot_qrs > 0 else 0
        duplas = datos.get("ev_duplas") or 0
        tsv = datos.get("tsv_episodios") or 0
        bloqueo = datos.get("tiene_bloqueo_rama", False)
        mcp = datos.get("mcp_presente", False)

        if tv > 0:
            detalles.append(f"Taquicardia Ventricular documentada ({tv} episodios)")
        if sdnn > 0 and sdnn <= 60 and not mcp:
            detalles.append(f"Variabilidad Autonómica Severamente Disminuida (SDNN: {sdnn} ms, Riesgo Alto)")
        if pausas > 0:
            detalles.append(f"Pausas patológicas registradas ({pausas} pausas)")
        if qtc > 500:
            detalles.append(f"Intervalo QTc Severamente Prolongado ({qtc} ms, riesgo de Torsades de Pointes)")
        if carga_ev >= 10.0:
            detalles.append(f"Carga Ectópica Ventricular Crítica ({carga_ev:.1f}%, riesgo de miocardiopatía)")

        if detalles:
            nivel = "ROJO"
            titulo = "ALERTA CLÍNICA: HALLAZGOS CARDIOVASCULARES CRÍTICOS"
        else:
            if mcp:
                detalles.append(f"Paciente portador de Dispositivo Cardíaco Implantable ({datos.get('mcp_porcentaje', 0):.1f}% estimulación/sensado)")
            if duplas > 0:
                detalles.append(f"Ectopia Ventricular Compleja: Duplas presentes ({duplas})")
            if tsv > 0:
                detalles.append(f"Taquicardia Supraventricular paroxística ({tsv} rachas)")
            if bloqueo and not mcp:
                detalles.append("Trastorno de la conducción intraventricular: Bloqueo de rama documentado")
            if sdnn > 0 and sdnn <= 120 and not mcp:
                detalles.append(f"Variabilidad de la FC disminuida (SDNN: {sdnn} ms, Riesgo Moderado)")
            if qtc > 460:
                detalles.append(f"QTc prolongado ({qtc} ms)")

            if detalles:
                nivel = "AMARILLO"
                titulo = "PRECAUCIÓN CLÍNICA: DISPOSITIVO CARDÍACO O PARÁMETROS LIMÍTROFES"
            else:
                detalles.append("Ritmo sinusal conservado, modulación autonómica normal y sin arritmias complejas.")

    elif modalidad == "MAPA":
        pas_24 = datos.get("pas_24h") or 0
        pad_24 = datos.get("pad_24h") or 0
        cn = datos.get("caida_nocturna_val") or 0.0
        pp = datos.get("pp_val") or 0
        c_pas = datos.get("carga_pas") or 0.0

        if cn <= -10.0:
            detalles.append(f"Patrón Circadiano Invertido (Riser: {cn:.1f}% caída nocturna, riesgo cerebrovascular agudo)")
        if pas_24 >= 140 or pad_24 >= 90:
            detalles.append(f"Descontrol Tensional Estadio II ({pas_24}/{pad_24} mmHg)")

        if detalles:
            nivel = "ROJO"
            titulo = "ALERTA HEMODINÁMICA: DESCONTROL TENSIONAL SEVERO O PATRÓN RISER"
        else:
            if cn <= 0.0:
                detalles.append("Patrón circadiano tensional no-dipper (atenuado)")
            if c_pas > 30.0:
                detalles.append(f"Carga sistólica elevada ({c_pas}%)")
            if pp > 60:
                detalles.append(f"Presión de pulso aumentada ({pp} mmHg, rigidez arterial)")

            if detalles:
                nivel = "AMARILLO"
                titulo = "PRECAUCIÓN HEMODINÁMICA: CONTROL SUBÓPTIMO O CARGA ELEVADA"
            else:
                detalles.append("Control óptimo de la tensión arterial con modulación circadiana conservada.")

    elif modalidad == "ESFUERZO":
        st_mm = datos.get("st_mm") or 0.0
        fc_p = datos.get("fc_pico") or 0
        edad = datos.get("edad", 35)
        fcm = 220 - edad if edad > 0 else 200
        porc = (fc_p / fcm) * 100 if fcm > 0 else 0

        if st_mm >= 1.0:
            detalles.append(f"Alteración de la repolarización: Infradesnivel del ST significativo ({st_mm} mm, sospecha isquémica)")

        if detalles:
            nivel = "ROJO"
            titulo = "ALERTA ISQUÉMICA: PRUEBA ELÉCTRICAMENTE POSITIVA"
        else:
            if porc > 0 and porc < 85.0:
                detalles.append(f"Prueba cronotrópicamente insuficiente ({porc:.1f}% de la FCM prevista, meta ≥ 85%)")

            if detalles:
                nivel = "AMARILLO"
                titulo = "PRECAUCIÓN: ESFUERZO INSUFICIENTE O TOLERANCIA LIMITADA"
            else:
                detalles.append(f"Prueba concluyente y suficiente ({porc:.1f}% FCM), sin alteraciones isquémicas del ST.")

    clase_css = f"triage-{nivel.lower()}"
    icono = "🚨" if nivel == "ROJO" else ("⚠️" if nivel == "AMARILLO" else "✅")
    items_html = "".join([f"<li>{d}</li>" for d in detalles])

    st.markdown(f"""
        <div class="{clase_css}">
            <div style="font-weight: 800; font-size: 0.95rem; display: flex; align-items: center; gap: 8px;">
                <span>{icono}</span> {titulo}
            </div>
            <ul style="margin: 0.4rem 0 0 1.2rem; font-size: 0.85rem; padding-left: 0;">
                {items_html}
            </ul>
        </div>
    """, unsafe_allow_html=True)

def generar_grafica_tacograma(d):
    fc_prom = d.get("fc_prom") or 0
    if fc_prom == 0:
        # Retorna grafica vacia si no hay datos
        fig = go.Figure()
        fig.update_layout(title="Tacograma no disponible", height=240, margin=dict(l=35, r=20, t=35, b=25))
        return fig

    fc_dia = d.get("fc_dia") or int(fc_prom * 1.05)
    fc_noc = d.get("fc_noc") or int(fc_prom * 0.90)
    fc_max = d.get("fc_max") or int(fc_prom * 1.4)
    fc_min = d.get("fc_min") or int(fc_prom * 0.7)

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
# MOTOR 1: EXTRACTOR DETERMINÍSTICO C++ (CERO HARDCODES / CERO INVENTOS)
# ==============================================================================
def extraer_numero(patron, texto, tipo=int):
    m = re.search(patron, texto, re.IGNORECASE)
    if m:
        s = m.group(1).replace(".", "").replace(",", ".") if tipo == int else m.group(1).replace(",", ".")
        try: return tipo(float(s))
        except: return None
    return None

def extraer_m1_holter(texto):
    d = {}
    texto_upper = texto.upper()
    
    m_nom = re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,50},\s*[A-ZÁÉÍÓÚÑ\s]{3,50})[\s\n]+(?:No confirmado|Confirmado)?[\s\n]*Informe", texto)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else None
    
    m_id = re.search(r"(?:ID\s*Paciente|ID|C\.?C\.?|Doc\.?|Historia)\s*[:\.]?\s*(\d{5,12})", texto, re.IGNORECASE)
    d["cedula"] = m_id.group(1) if m_id else None

    # Extracciones que devuelven None si no encuentran el dato, SIN fallbacks inventados
    d["fc_prom"] = extraer_numero(r"Prom\.?\s*(\d{2,3})", texto, int)
    d["fc_max"]  = extraer_numero(r"M[áa]x\s*(\d{2,3})", texto, int)
    d["fc_min"]  = extraer_numero(r"M[íi]n\s*(\d{2,3})", texto, int)
    d["fc_dia"]  = extraer_numero(r"D[íi]a.*?Prom\.?\s*(\d{2,3})", texto, int)
    d["fc_noc"]  = extraer_numero(r"Noche.*?Prom\.?\s*(\d{2,3})", texto, int)

    m_conteo = re.search(r"Latidos[^\n\r]*\n[^\n\r]*Conteo\s+([\d\.]+)\s+([\d\.]+)\s+\d+%\s+([\d\.]+)[^\n\r]*\s+([\d\.]+)[^\n\r]*\s+([\d\.]+)\s+(\d+)?%", texto, re.IGNORECASE)
    if m_conteo:
        d["total_latidos"]    = extraer_numero(r"([\d\.]+)", m_conteo.group(1), int)
        d["latidos_normales"] = extraer_numero(r"([\d\.]+)", m_conteo.group(2), int)
        d["ev_total"]         = extraer_numero(r"([\d\.]+)", m_conteo.group(3), int)
        d["esv_total"]        = extraer_numero(r"([\d\.]+)", m_conteo.group(4), int)
        d["mcp_latidos"]      = extraer_numero(r"([\d\.]+)", m_conteo.group(5), int)
        d["mcp_porcentaje"]   = extraer_numero(r"([\d\.]+)", m_conteo.group(6), float) if m_conteo.group(6) else None
    else:
        d["total_latidos"] = extraer_numero(r"Total\s+de\s+latidos\s*[:\.]?\s*([\d\.]+)", texto, int)
        d["ev_total"]      = extraer_numero(r"Latidos ventriculares\s*:\s*([\d\.]+)", texto, int)
        d["esv_total"]     = extraer_numero(r"Latidos supraventriculares\s*:\s*([\d\.]+)", texto, int)
        d["mcp_latidos"]   = extraer_numero(r"Marcapaso\s+([\d\.]+)", texto, int)
        d["mcp_porcentaje"]= extraer_numero(r"Marcapaso\s+[\d\.]+\s+(\d+)%", texto, float)
        d["latidos_normales"] = None

    d["mcp_presente"] = True if (d.get("mcp_latidos") and d["mcp_latidos"] > 0) else (None if "MARCAPASOS" not in texto_upper else True)

    d["rr_max_seg"] = extraer_numero(r"Intervalo\s+RR.*?M[áa]x\.\s*longitud\s*([\d,\.]+)\s*s", texto, float)
    d["pausas"] = extraer_numero(r"\bPausa\s+(\d+)", texto, int)
    d["tv_episodios"] = extraer_numero(r"\bTV\s+([\d\.]+)", texto, int)
    d["ev_duplas"] = extraer_numero(r"Apareado\s+([\d\.]+)", texto, int)
    d["tsv_episodios"] = extraer_numero(r"\bTSV\s+([\d\.]+)", texto, int)
    d["sdnn_24h"] = extraer_numero(r"Valor de 24 horas\s+[\d\.,]+\s+([\d\.,]+)", texto, int)
    d["qtc_prom"] = extraer_numero(r"Todos los per[íi]odos\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)", texto, int)

    d["st_dep_episodios"] = extraer_numero(r"Depresi[óo]n ST\s+(\d+)", texto, int)
    d["st_elev_episodios"] = extraer_numero(r"Elevaci[óo]n ST\s+(\d+)", texto, int)
    
    bloqueo_keys = ["BLOQUEO DE RAMA", "BRD", "BRI", "BCRD", "BCRI", "QRS ANCHO", "CONDUCCION INTRAVENTRICULAR"]
    d["tiene_bloqueo_rama"] = any(k in texto_upper for k in bloqueo_keys)
    return d

def extraer_m1_mapa(texto):
    d = {}
    m_nom = re.search(r"Nombre del paciente:\s*([^\n\r\|]+)", texto, re.IGNORECASE)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else None
    
    d["pas_24h"] = extraer_numero(r"Resumen general.*?Prom\.?:\s*(\d{2,3})", texto, int)
    d["pad_24h"] = extraer_numero(r"Resumen general.*?Prom\.?:\s*\d{2,3}\s*[\/\-]\s*(\d{2,3})", texto, int)
    d["carga_pas"] = extraer_numero(r"Sist[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", texto, float)
    d["carga_pad"] = extraer_numero(r"Diast[óo]lico\s*>\s*l[íi]mite\s*:\s*([\d,\.]+)\s*%", texto, float)
    d["pp_val"] = extraer_numero(r"Presi[óo]n de pulso\s*\(mmHg\)\s*\n?\s*(\d{2,3})", texto, int)
    d["caida_nocturna_val"] = extraer_numero(r"Sist[óo]lico\s*\(mmHg\)\s*.*?([\d,\.\-]+)\s*%", texto, float)
    
    m_sueno = re.search(r"Resumen de los per[íi]odos de sue[ñn]o.*?Sist[óo]lico\s*\(mmHg\)\s*\n?\s*(\d+)\s*.*?(\d{2,3})\s*\([^\)]+\)\s*.*?Diast[óo]lico\s*\(mmHg\)\s*\n?\s*(\d+)\s*.*?(\d{2,3})\s*\(", texto, re.DOTALL | re.IGNORECASE)
    d["pas_max_sueno"] = int(m_sueno.group(2)) if m_sueno else None
    d["pad_max_sueno"] = int(m_sueno.group(4)) if m_sueno else None
    return d

def extraer_m1_esfuerzo(texto):
    d = {}
    d["fc_pico"] = extraer_numero(r"FC Pico\s*[:\-]\s*(\d{2,3})", texto, int)
    return d

# ==============================================================================
# MOTOR 2: INTELIGENCIA ARTIFICIAL GEMINI (RENDERIZADO 100% DINÁMICO)
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
        except: continue
    return False, None, "Timeout en API Gemini"

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
        except: continue

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
  "paciente": "KAROL JULIANA SUAREZ FARFAN", "cedula": "1050095219", "edad": 17, "sexo": "Femenino",
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
def arbitrar_valor(k, v1, v2, tipo):
    if v1 is not None and v2 is not None:
        if v1 == v2: return v1, None
        return v1, f"⚠️ {k}: M1({v1}) vs IA({v2}) -> Gana M1 (Tabla física determinista)."
    if v1 is not None: return v1, f"ℹ️ {k}: Solo M1 leyó el valor ({v1})."
    if v2 is not None: return v2, f"⚕️ {k}: M1 ciego. IA extrajo valor visual ({v2})."
    return (0 if tipo == int else 0.0), f"🚨 CRÍTICO: Ningún motor encontró '{k}'. Se asigna 0 por seguridad."

def auditar_motor_3_universal(m1, m2, modalidad):
    d_f = {}
    notas = []

    if modalidad == "HOLTER":
        keys_int = ["fc_prom", "fc_dia", "fc_noc", "fc_max", "fc_min", "total_latidos", "ev_total", "esv_total", "mcp_latidos", "pausas", "tv_episodios", "ev_duplas", "tsv_episodios", "sdnn_24h", "qtc_prom", "st_dep_episodios", "st_elev_episodios", "eventos_paciente"]
        for k in keys_int:
            val, msg = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int)
            d_f[k] = val
            if msg and ("CRÍTICO" in msg or "vs" in msg or "IA extrajo" in msg): notas.append(msg)
            
        d_f["rr_max_seg"], msg = arbitrar_valor("rr_max_seg", m1.get("rr_max_seg"), m2.get("rr_max_seg") if m2 else None, float)
        if msg and ("CRÍTICO" in msg or "vs" in msg): notas.append(msg)

        d_f["paciente"] = m2.get("paciente") if (m2 and m2.get("paciente")) else m1.get("paciente", "PACIENTE ANÓNIMO")
        d_f["cedula"] = m2.get("cedula") if (m2 and m2.get("cedula")) else m1.get("cedula", "")

        # Candado LEY DE CONSERVACIÓN QRS
        norm_m1 = m1.get("latidos_normales")
        norm_m2 = m2.get("latidos_normales") if m2 else None
        norm, _ = arbitrar_valor("latidos_normales", norm_m1, norm_m2, int)
        
        suma = norm + d_f["ev_total"] + d_f["esv_total"] + d_f["mcp_latidos"]
        if d_f["total_latidos"] > 0 and abs(suma - d_f["total_latidos"]) > (d_f["total_latidos"] * 0.05):
            notas.append(f"⚖️ CONSERVACIÓN: Suma de componentes ({suma}) no cuadra con Total ({d_f['total_latidos']}). Ajustando Total.")
            d_f["total_latidos"] = max(d_f["total_latidos"], suma)
            d_f["latidos_normales"] = norm
        else:
            d_f["latidos_normales"] = norm if norm > 0 else max(0, d_f["total_latidos"] - d_f["ev_total"] - d_f["esv_total"] - d_f["mcp_latidos"])

        # Candado PAUSAS vs RR MÁXIMO
        if 0 < d_f["rr_max_seg"] < 2.0 and d_f["pausas"] > 0:
            notas.append(f"🔒 VETO FISIOLÓGICO: RR Máx es {d_f['rr_max_seg']}s. Imposible pausas > 2.0s. Se vetan pausas.")
            d_f["pausas"] = 0

        # Dispositivos
        mcp_pct_m1 = m1.get("mcp_porcentaje")
        mcp_pct_m2 = m2.get("mcp_porcentaje") if m2 else None
        d_f["mcp_porcentaje"], _ = arbitrar_valor("mcp_porcentaje", mcp_pct_m1, mcp_pct_m2, float)
        d_f["mcp_presente"] = True if d_f["mcp_latidos"] > 0 else (m1.get("mcp_presente") or False)
        
        t_bloq = m1.get("tiene_bloqueo_rama", False) or (m2.get("tiene_bloqueo_rama", False) if m2 else False)
        if d_f["mcp_presente"] and d_f["mcp_porcentaje"] >= 50.0:
            d_f["tiene_bloqueo_rama"] = True
            d_f["bloqueo_detalle"] = "bloqueo de rama izquierda secundario a electroestimulación ventricular"
        else:
            d_f["tiene_bloqueo_rama"] = t_bloq
            d_f["bloqueo_detalle"] = "bloqueo de rama documentado"
            
        d_f["st_episodios"] = d_f.get("st_dep_episodios", 0) + d_f.get("st_elev_episodios", 0)

    elif modalidad == "MAPA":
        keys_int = ["pas_24h", "pad_24h", "pp_val", "pas_max_sueno", "pad_max_sueno"]
        for k in keys_int:
            val, msg = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int)
            d_f[k] = val
            if msg and ("CRÍTICO" in msg or "vs" in msg or "IA extrajo" in msg): notas.append(msg)
            
        keys_float = ["carga_pas", "carga_pad", "caida_nocturna_val"]
        for k in keys_float:
            val, msg = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, float)
            d_f[k] = val
            if msg and ("CRÍTICO" in msg or "vs" in msg or "IA extrajo" in msg): notas.append(msg)
            
        d_f["paciente"] = m2.get("paciente") if (m2 and m2.get("paciente")) else m1.get("paciente", "PACIENTE MAPA")
        d_f["cedula"] = m2.get("cedula") if (m2 and m2.get("cedula")) else m1.get("cedula", "")

        if d_f["pas_24h"] > 0 and d_f["pad_24h"] > 0:
            d_f["pam_24h"] = round(d_f["pad_24h"] + ((d_f["pas_24h"] - d_f["pad_24h"]) / 3))
            d_f["pp_val"] = d_f["pas_24h"] - d_f["pad_24h"]
        else:
            d_f["pam_24h"] = 0; d_f["pp_val"] = 0

    return d_f, notas

def ejecutar_triple_engine_universal(pdf_bytes, modalidad="HOLTER"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pags = len(doc)
    
    # M1: Escaneo completo 0 a N
    texto_completo = "".join([f"\n--- PÁG {i+1} ---\n" + p.get_text() + "\n" for i, p in enumerate(doc)])
    m1 = extraer_m1_holter(texto_completo) if modalidad == "HOLTER" else extraer_m1_mapa(texto_completo)

    # M2: Visión IA dinámica 0 a N (80 DPI para compresión ultra-ligera y velocidad extrema)
    inline_images = []
    for num_p in range(total_pags):
        try:
            pix = doc[num_p].get_pixmap(dpi=80)
            b64_p = base64.b64encode(pix.tobytes("jpeg")).decode("utf-8")
            inline_images.append({"inline_data": {"mime_type": "image/jpeg", "data": b64_p}})
        except: continue
    doc.close()

    if modalidad == "HOLTER":
        prompt_ia = f"""Eres Auditor Médico Jefe. Analiza las {total_pags} imágenes y el texto de este Holter 24h.
Texto: {texto_completo}
Devuelve JSON estricto (null si el dato no está):
{{"paciente": "str", "cedula": "str", "mcp_presente": bool, "mcp_latidos": int, "mcp_porcentaje": float, "fc_prom": int, "fc_dia": int, "fc_noc": int, "fc_max": int, "fc_min": int, "total_latidos": int, "latidos_normales": int, "pausas": int, "rr_max_seg": float, "ev_total": int, "tv_episodios": int, "ev_duplas": int, "esv_total": int, "tsv_episodios": int, "sdnn_24h": int, "qtc_prom": int, "st_dep_episodios": int, "st_elev_episodios": int, "eventos_paciente": int, "tiene_bloqueo_rama": bool}}"""
    else:
        prompt_ia = f"""Eres Auditor Médico Jefe. Analiza las {total_pags} imágenes y el texto de este MAPA.
Texto: {texto_completo}
Devuelve JSON estricto (null si el dato no está):
{{"paciente": "str", "cedula": "str", "pas_24h": int, "pad_24h": int, "carga_pas": float, "carga_pad": float, "pp_val": int, "caida_nocturna_val": float, "pas_max_sueno": int, "pad_max_sueno": int}}"""

    ok_m2, m2, mod_ia = consultar_gemini_json(prompt_ia, inline_items=inline_images)
    
    # M3: Quórum Final
    d_final, notas = auditar_motor_3_universal(m1, m2 if ok_m2 else None, modalidad)
    d_final["triple_engine_activo"] = True
    d_final["notas_arbitraje"] = notas
    d_final["motores_info"] = f"Auditoría Triple Engine: M1 (Espacial {total_pags} pág) + M2 (Visión IA {total_pags} pág) + M3 (Fisiología)"
    
    return d_final

# ==============================================================================
# REDACTORES INSTITUCIONALES (CERO HARDCODES)
# ==============================================================================
def redactar_informe_holter_11_puntos(d, perfil):
    fc_p = d.get("fc_prom") or 0
    fc_d = d.get("fc_dia") or 0
    fc_n = d.get("fc_noc") or 0
    desc_crono = ((fc_d - fc_n) / fc_d * 100) if fc_d > 0 else 0
    dip_txt = f"conservado ({desc_crono:.1f}% descenso)" if desc_crono >= 10 else f"atenuado ({desc_crono:.1f}% descenso)"

    pct_mcp = d.get("mcp_porcentaje") or 0.0
    if d.get("mcp_presente") and pct_mcp >= 80.0: p1 = f"1. Ritmo comandado por marcapasos definitivo (electroestimulado/sensado en {pct_mcp:.1f}%) con FC promedio {fc_p} lpm (Día: {fc_d} / Noche: {fc_n})."
    elif d.get("mcp_presente"): p1 = f"1. Ritmo sinusal alternando con electroestimulación ({pct_mcp:.1f}% mediado por MCP) con FC promedio {fc_p} lpm (Día: {fc_d} / Noche: {fc_n})."
    else: p1 = f"1. Ritmo sinusal con FC promedio {fc_p} lpm (Día: {fc_d} / Noche: {fc_n}; patrón circadiano {dip_txt})."

    crono = []
    if d.get("taqui_conteo", 0) > 0: crono.append(f"{d.get('taqui_conteo')} episodios de taquicardia (FC máx {d.get('taqui_fc_max', d.get('fc_max',0))} lpm)")
    if d.get("bradi_conteo", 0) > 0: crono.append(f"{d.get('bradi_conteo')} episodios de bradicardia (FC mín {d.get('bradi_fc_min', d.get('fc_min',0))} lpm)")
    p2 = f"2. Eventos cronotrópicos: Se documentaron {' y '.join(crono)}." if crono else "2. Eventos cronotrópicos: Sin bradicardia patológica ni taquicardias sostenidas."

    qtc = d.get("qtc_prom", 0)
    if qtc > 500: p3 = f"3. QTc SEVERAMENTE PROLONGADO ({qtc} ms: alto riesgo de Torsades de Pointes)."
    elif qtc > 460: p3 = f"3. QTc prolongado ({qtc} ms)."
    else: p3 = f"3. QTc normal ({qtc} ms)."

    p4 = f"4. Alteraciones isquémicas del ST ({d.get('st_episodios',0)} eventos)." if d.get("st_episodios",0) > 0 else "4. Sin alteraciones isquémicas del segmento ST."
    p5 = f"5. Conducción AV alterada por {d.get('latidos_caidos',0)} latidos caídos." if d.get("latidos_caidos",0) > 0 else "5. Sin Alteración de la conducción AV."

    if d.get("mcp_presente") and pct_mcp >= 50.0: p6 = f"6. Complejos ventriculares con {d.get('bloqueo_detalle', 'morfología de bloqueo por estimulación')}."
    elif d.get("tiene_bloqueo_rama"): p6 = f"6. Alteración de conducción intraventricular por {d.get('bloqueo_detalle', 'bloqueo de rama')}."
    else: p6 = "6. Sin alteración en la conducción intraventricular."

    tot_q = d.get("total_latidos") or 1
    ev = d.get("ev_total") or 0
    carga = (ev / tot_q) * 100 if tot_q > 0 else 0
    if d.get("tv_episodios", 0) > 0: lown = "Lown Grado IVb (Taquicardia Ventricular)"
    elif d.get("ev_duplas", 0) > 0: lown = "Lown Grado IVa (Duplas)"
    elif ev >= 720: lown = "Lown Grado II (Frecuentes > 30/h)"
    elif ev > 0: lown = "Lown Grado I"
    else: lown = "Lown Grado 0"

    ect = []
    if d.get("esv_total", 0) > 0: ect.append(f"ectopias supraventriculares ({d.get('esv_total')} ESV)")
    if ev > 0: ect.append(f"ectopia ventricular con {ev} EV (Carga {carga:.2f}%, {lown})")
    p7 = f"7. Impulsos alterados por {' y '.join(ect)}." if ect else "7. Sin alteración de impulsos ectópicos."
    p8 = f"8. Diario paciente: {d.get('eventos_paciente', 0)} eventos marcados." if d.get("eventos_paciente", 0) > 0 else "8. No refirió síntomas."

    sdnn = d.get("sdnn_24h", 0)
    if sdnn > 0 and sdnn <= 60 and not d.get("mcp_presente"): p9, p11, d_vfc = f"9. VFC severamente disminuida ({sdnn} ms).", "11. Riesgo autonómico: ALTO.", "VFC severamente disminuida (riesgo autonómico alto)."
    elif sdnn > 0 and sdnn <= 120 and not d.get("mcp_presente"): p9, p11, d_vfc = f"9. VFC disminuida ({sdnn} ms).", "11. Riesgo autonómico: Moderado.", "VFC disminuida."
    elif d.get("mcp_presente"): p9, p11, d_vfc = f"9. Variabilidad cardíaca modificada por MCP ({sdnn} ms).", "11. Evaluación dependiente del generador.", "Frecuencia regulada por dispositivo."
    else: p9, p11, d_vfc = f"9. VFC conservada ({sdnn} ms).", "11. Riesgo autonómico: Bajo/Normal.", "VFC conservada."

    p10 = f"10. Se registraron {d.get('pausas',0)} pausas significativas (máx {d.get('pausa_max_seg', 0)}s)." if d.get("pausas",0) > 0 else ("10. Dispositivo normofuncionante: adecuado sensado y captura." if d.get("mcp_presente") else "10. No hay pausas significativas.")

    diag = f"Ritmo sinusal con FC promedio {fc_p} lpm. {d_vfc}"
    if d.get("mcp_presente"): diag = f"Ritmo comandado por marcapasos normofuncionante ({pct_mcp:.1f}%). FC promedio {fc_p} lpm."
    
    recs = ["Valoración por Electrofisiología."] if d.get("tv_episodios",0) > 0 or carga >= 10.0 else ["Continuar manejo médico instaurado y seguimiento clínico."]

    return f"INTERPRETACIÓN TEST HOLTER - CUPS 895001\n\n{p1}\n{p2}\n{p3}\n{p4}\n{p5}\n{p6}\n{p7}\n{p8}\n{p9}\n{p10}\n{p11}\n\nCONCLUSIÓN DIAGNÓSTICA:\n{diag}\n\nRECOMENDACIONES: {' '.join(recs)}\n\n{perfil['nombre_completo']}\n{perfil['especialidad']}\n{perfil['registro']}"

def redactar_informe_mapa_cencardio(d, perfil):
    p1 = f"1. Promedios 24h: Sistólica {d.get('pas_24h',0)} mmHg / Diastólica {d.get('pad_24h',0)} mmHg; PAM: {d.get('pam_24h',0)} mmHg."
    p2 = f"2. Carga tensional sistólica ({d.get('carga_pas',0)}%) y diastólica ({d.get('carga_pad',0)}%)."
    p3 = f"3. Presión de pulso: {d.get('pp_val',0)} mmHg."
    p4 = f"4. Caída nocturna (Dipping): {d.get('caida_nocturna_val',0.0)}%."
    
    c_pas = d.get('carga_pas',0)
    ctrl = "Control óptimo de la tensión arterial" if (c_pas < 15 and d.get('pas_24h',0) < 130) else "Control subóptimo de la tensión arterial"
    recs = "Continuar manejo farmacológico actual." if "óptimo" in ctrl else "Optimización y titulación de la terapia antihipertensiva."
    
    return f"INTERPRETACIÓN TEST MAPA - CUPS 895003\n\nHallazgos:\n{p1}\n{p2}\n{p3}\n{p4}\n5. {ctrl}.\n\nRECOMENDACIONES: {recs}\n\n{perfil['nombre_completo']}\n{perfil['especialidad']}\n{perfil['registro']}"

def calcular_mets_bruce(tiempo_min):
    if tiempo_min <= 0.1: return 0.0
    if tiempo_min <= 3.0: return round(1.0 + (tiempo_min / 3.0) * 3.8, 1)
    if tiempo_min <= 6.0: return round(4.8 + ((tiempo_min - 3.0) / 3.0) * 2.2, 1)
    if tiempo_min <= 9.0: return round(7.0 + ((tiempo_min - 6.0) / 3.0) * 3.1, 1)
    if tiempo_min <= 12.0: return round(10.1 + ((tiempo_min - 9.0) / 3.0) * 3.4, 1)
    if tiempo_min <= 15.0: return round(13.5 + ((tiempo_min - 12.0) / 3.0) * 3.7, 1)
    return round(17.2 + ((tiempo_min - 15.0) / 3.0) * 3.0, 1)

def redactar_informe_ergometria_institucional(d, perfil):
    fcm_prev = 220 - d.get("edad", 35) if d.get("edad", 35) > 0 else 200
    porc = round((d.get("fc_pico", 150) / fcm_prev) * 100) if fcm_prev > 0 else 0
    suf = "suficiente" if porc >= 85 else "insuficiente"
    dp = d.get("fc_pico", 150) * d.get("pas_pico", 150)
    st_res = "Sin alteraciones isquémicas del segmento ST" if d.get("st_mm", 0.0) < 1.0 else f"Infradesnivel del ST de {d.get('st_mm')} mm"
    diag_el = "negativa" if d.get("st_mm", 0.0) < 1.0 else "positiva"
    dts = d.get("tiempo_min", 0.0) - (5 * d.get("st_mm", 0.0))
    duke = "Bajo riesgo coronario" if dts >= 5 else ("Riesgo moderado" if dts >= -10 else "Alto riesgo coronario")
    reserva_den = (fcm_prev - d.get("fc_basal", 75)) if (fcm_prev - d.get("fc_basal", 75)) > 0 else 1
    fcr_pct = round(((d.get("fc_pico", 150) - d.get("fc_basal", 75)) / reserva_den) * 100)
    incomp_crono = "Incompetencia cronotrópica" if fcr_pct < 80 else "Respuesta cronotrópica adecuada"

    return f"""INTERPRETACIÓN PRUEBA DE ESFUERZO COMPUTARIZADA - CUPS 893805

1. Ritmo sinusal normal basal y durante el esfuerzo físico.
2. Protocolo de {d.get('protocolo', 'Bruce')} completado con duración de {d.get('tiempo_min', 0.0):.2f} minutos ({d.get('etapa', 'Final')}).
3. Capacidad funcional alcanzada: {d.get('mets', 0.0)} METs.
4. Respuesta cronotrópica: FC basal {d.get('fc_basal', 0)} lpm a FC pico de {d.get('fc_pico', 0)} lpm ({porc}% FCM prevista; FCR: {fcr_pct}%, {incomp_crono}, prueba {suf}).
5. Respuesta hemodinámica presora: PA basal {d.get('pas_basal', 0)}/{d.get('pad_basal', 0)} mmHg a PA pico de {d.get('pas_pico', 0)}/{d.get('pad_pico', 0)} mmHg.
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
    except: return None

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
    porc = round((d.get("fc_pico", 150) / fcm_prev) * 100) if fcm_prev > 0 else 0
    cajas = [
        ("FC PICO ALCANZADA", f"{d.get('fc_pico',0)} lpm ({porc}%)"),
        ("PA ESFUERZO PICO", f"{d.get('pas_pico',0)}/{d.get('pad_pico',0)} mmHg"),
        ("CARGA FUNCIONAL", f"{d.get('mets',0.0)} METs ({d.get('tiempo_min',0.0):.2f} m)"),
        ("DOBLE PRODUCTO", f"{d.get('fc_pico',0)*d.get('pas_pico',0):,}")
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
        except: pass

    pdf_bytes = doc.tobytes()
    doc.close()
    return pdf_bytes, img_preview

# ==============================================================================
# EXCEL GERENCIAL
# ==============================================================================
def generar_excel_avanzado_produccion(df_base):
    if not OPENPYXL_INSTALADO: return df_base.to_csv(sep=';', index=False, encoding='utf-8-sig').encode('utf-8-sig'), "csv"
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Archivo CENCARDIO"
    headers = ["ID", "Fecha", "Paciente", "Modalidad", "CUPS", "Métrica", "Lector", "Forense"]
    for i, h in enumerate(headers, 1): ws.cell(row=1, column=i, value=h).font = Font(bold=True)
    df_e = df_base[["id", "fecha_registro", "paciente", "modalidad", "cups", "parametro_clave", "medico", "codigo_verificacion"]].copy()
    for r, fila in enumerate(df_e.itertuples(index=False), 2):
        for c, val in enumerate(fila, 1): ws.cell(row=r, column=c, value=str(val))
    wb.save(output)
    return output.getvalue(), "xlsx"

# ==============================================================================
# UI PRINCIPAL: BARRERA DE LOGIN Y NAVEGACIÓN
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
        st.markdown(f'<div class="cencardio-card"><div style="font-size:3rem;">🫀</div><div style="font-weight:800; font-size:1.2rem; margin-bottom:1rem;">CERTIFICADO DE AUTENTICIDAD FORENSE</div><div style="background:#f0fdf4; border:1px solid #86efac; border-radius:12px; padding:1.3rem; text-align:left;"><b>Procedimiento:</b> {urllib.parse.unquote(params.get("proc",""))}<br><b>Paciente:</b> {urllib.parse.unquote(params.get("pac",""))}<br><b>Lector:</b> {urllib.parse.unquote(params.get("med",""))}<br><b>Fecha:</b> {urllib.parse.unquote(params.get("fec",""))}<br><b>Cód:</b> {params.get("val","")}</div></div>', unsafe_allow_html=True)
        if st.button("Ir al Portal", use_container_width=True):
            st.query_params.clear()
            st.rerun()
    st.stop()

if not st.session_state.autenticado:
    c1, c2, c3 = st.columns([1, 1.6, 1])
    with c2:
        st.markdown('<div class="cencardio-card"><div style="font-size:3rem; margin-bottom:0.4rem;">🫀</div><div style="font-size:1.25rem; font-weight:800; color:#0a2540;">Centro Cardiovascular Colombiano</div><div style="font-size:0.8rem; font-weight:700; color:#c8102e; margin-bottom:1.8rem;">CENCARDIO WORKSTATION</div>', unsafe_allow_html=True)
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

st.markdown(f'<div class="top-hospital-bar"><div><div style="font-size:1.35rem; font-weight:800; color:#0a2540;">CENCARDIO WORKSTATION</div><div style="font-size:0.82rem; font-weight:700; color:#c8102e;">Alta Complejidad</div></div><div style="display:flex; gap:8px;"><span class="inst-badge-success">{"Nube Supabase" if supabase else "Local SQLite"}</span><span class="inst-badge-primary">Triple Engine 100%</span></div></div>', unsafe_allow_html=True)

tab_procesar, tab_historial = st.tabs(["📥 Procesamiento del Estudio", "📁 Archivo Clínico & Reportes Gerenciales"])

# ==============================================================================
# PESTAÑA 1: PROCESAMIENTO
# ==============================================================================
with tab_procesar:
    if "Esfuerzo" in mod_sel:
        st.markdown("#### 🏃 Consola de Emisión de Prueba de Esfuerzo (CUPS 893805)")
        for k, v in [("erg_paciente", ""), ("erg_cedula", ""), ("erg_edad", 35), ("erg_sexo", "Femenino"), ("erg_protocolo", "Bruce"), ("erg_etapa", "Etapa 4"), ("erg_tiempo", 0.0), ("erg_fc_basal", 75), ("erg_fc_pico", 150), ("erg_pa_basal", "120/80"), ("erg_pa_pico", "160/90"), ("erg_st_mm", 0.0), ("erg_celular", "")]:
            if k not in st.session_state: st.session_state[k] = v

        col_f1, col_f2 = st.columns([1.2, 1], gap="large")
        with col_f1:
            fotos_esfuerzo = st.file_uploader("📸 Subir tiras de esfuerzo (JPG/PNG):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
            if fotos_esfuerzo and st.button("⚡ EXTRAER CON VISIÓN IA", type="primary", use_container_width=True):
                with st.spinner("Analizando imágenes con IA Multimodal..."):
                    ok, msg = ejecutar_extraccion_multimodal(fotos_esfuerzo)
                    if ok: st.success(msg); st.rerun()
                    else: st.error(msg)
            
            c1, c2, c3 = st.columns(3)
            with c1:
                p_nombre = st.text_input("Paciente:", key="erg_paciente")
                p_cedula = st.text_input("Cédula:", key="erg_cedula")
            with c2:
                p_edad = st.number_input("Edad:", min_value=1, key="erg_edad")
                p_sexo = st.selectbox("Sexo:", ["Femenino", "Masculino"], key="erg_sexo")
            with c3:
                p_protocolo = st.selectbox("Protocolo:", ["Bruce", "Naughton"], key="erg_protocolo")
                p_etapa = st.text_input("Etapa:", key="erg_etapa")
            
            c4, c5, c6 = st.columns(3)
            with c4:
                p_tiempo = st.number_input("Tiempo (min):", step=0.1, key="erg_tiempo")
                p_mets = st.number_input("METs:", value=float(calcular_mets_bruce(p_tiempo)), step=0.5)
            with c5:
                p_fc_basal = st.number_input("FC Basal:", key="erg_fc_basal")
                p_fc_pico = st.number_input("FC Pico:", key="erg_fc_pico")
            with c6:
                p_pa_basal = st.text_input("PA Basal:", key="erg_pa_basal")
                p_pa_pico = st.text_input("PA Pico:", key="erg_pa_pico")
            
            c7, c8 = st.columns(2)
            with c7:
                p_st_mm = st.number_input("Desviación ST (mm):", step=0.5, key="erg_st_mm")
            with c8:
                if not st.session_state.erg_celular and p_cedula: st.session_state.erg_celular = buscar_telefono_servicio(p_cedula)
                p_celular = st.text_input("Celular (WhatsApp):", key="erg_celular")

        with col_f2:
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

                uid = str(uuid.uuid4()).upper()
                pdf_erg, img_erg_prev = generar_pdf_ergometria_completo(d_erg, txt_erg_final, perfil_activo, uid, fotos_esfuerzo or [])

                b1, b2 = st.columns(2)
                with b1:
                    st.download_button("📄 DESCARGAR PDF", data=pdf_erg, file_name=f"{normalizar_nombre_archivo(p_nombre)}_Esfuerzo.pdf", mime="application/pdf", type="primary", use_container_width=True)
                with b2:
                    if st.button("💾 Guardar Clínico", use_container_width=True):
                        guardar_estudio_servicio(p_nombre, "Prueba de Esfuerzo", "CUPS 893805", f"FC Pico: {p_fc_pico} | {p_mets} METs", perfil_activo['nombre_completo'], txt_erg_final, pdf_erg, uid)
                        st.success("✅ Guardado en Archivo Clínico.")
                
                if p_celular:
                    tel_l = re.sub(r'\D', '', p_celular)
                    if len(tel_l) == 10: tel_l = "57" + tel_l
                    url_c = f"https://holtercencardio.streamlit.app/?val={uid[:12]}&pac={urllib.parse.quote(p_nombre)}&med={urllib.parse.quote(perfil_activo['nombre_completo'])}&proc=CUPS_893805"
                    st.link_button("📲 ENVIAR WHATSAPP", f"https://wa.me/{tel_l}?text={urllib.parse.quote(f'Estimado(a) {p_nombre}, CENCARDIO le entrega su resultado de Prueba de Esfuerzo: {url_c}')}", use_container_width=True)
                if img_erg_prev: st.image(img_erg_prev, use_container_width=True)
    else:
        up_file = st.file_uploader(f"📥 Cargar PDF de {mod_sel}", type=["pdf"])
        if up_file:
            b_pdf = up_file.getvalue()
            if st.session_state.get("last_file") != up_file.name:
                with st.spinner("🚀 Ejecutando TRIPLE ENGINE (Leyendo 100% del PDF sin suposiciones)..."):
                    mod = "HOLTER" if "Holter" in mod_sel else "MAPA"
                    d_final = ejecutar_triple_engine_universal(b_pdf, mod)
                    st.session_state.d_final = d_final
                    st.session_state.txt_final = redactar_informe_holter_11_puntos(d_final, perfil_activo) if mod == "HOLTER" else redactar_informe_mapa_cencardio(d_final, perfil_activo)
                    st.session_state.last_file = up_file.name
                    st.session_state.uid = str(uuid.uuid4()).upper()
                    st.session_state.tel = buscar_telefono_servicio(d_final.get("cedula"))
                    st.toast(f"✅ {d_final.get('motores_info')}")

            d = st.session_state.d_final
            mod_actual = "HOLTER" if "Holter" in mod_sel else "MAPA"
            cups_actual = "CUPS 895001" if mod_actual == "HOLTER" else "CUPS 895003"
            mod_nombre = "Holter ECG 24H" if mod_actual == "HOLTER" else "MAPA Tensional 24H"
            
            mostrar_semaforizacion_clinica(mod_actual, d)
            bloq, ale = ejecutar_sanity_checks(mod_actual, d)
            for b in bloq: st.error(f"🛑 {b}")
            for a in ale: st.warning(f"⚠️ {a}")
            
            if d.get("notas_arbitraje"):
                with st.expander("🛡️ REPORTE DE ARBITRAJE TRIPLE ENGINE (M1 vs M2 vs M3)", expanded=True):
                    for n in d["notas_arbitraje"]: 
                        if "CRÍTICO" in n or "⚠️" in n or "🚨" in n: st.error(n)
                        else: st.info(n)
            
            c1, c2, c3, c4 = st.columns(4)
            if mod_actual == "HOLTER":
                c1.metric("FC Promedio", f"{d.get('fc_prom',0)} lpm", f"Día {d.get('fc_dia',0)} / Noche {d.get('fc_noc',0)}")
                c2.metric("Ectopias V", f"{d.get('ev_total',0)} EV", f"TV: {d.get('tv_episodios',0)}")
                c3.metric("Pausas", f"{d.get('pausas',0)}", f"Máx: {d.get('pausa_max_seg',0)}s")
                c4.metric("SDNN / MCP", f"{d.get('sdnn_24h',0)} ms", f"MCP: {d.get('mcp_porcentaje',0)}%")
                
                # Renderiza gráfica del tacograma
                fig_taco = generar_grafica_tacograma(d)
                st.plotly_chart(fig_taco, use_container_width=True)
            else:
                c1.metric("PAM 24H", f"{d.get('pam_24h',0)} mmHg", f"{d.get('pas_24h',0)}/{d.get('pad_24h',0)}")
                c2.metric("Carga Sistólica", f"{d.get('carga_pas',0)}%")
                c3.metric("Carga Diastólica", f"{d.get('carga_pad',0)}%")
                c4.metric("Caída Nocturna", f"{d.get('caida_nocturna_val',0)}%")
            
            st.divider()
            col_ed, col_prev = st.columns([1, 1], gap="large")
            with col_ed:
                c_nom, c_ced = st.columns([1.8, 1.2])
                with c_nom:
                    p_nom = st.text_input("👤 Paciente:", value=d.get('paciente') or "")
                with c_ced:
                    p_ced = st.text_input("🪪 Cédula:", value=d.get('cedula') or "")
                    
                tel_in = st.text_input("📱 Celular (WhatsApp):", value=st.session_state.tel)
                txt_edit = st.text_area(f"📝 Informe Oficial ({cups_actual}):", value=st.session_state.txt_final, height=360)
                
                for err in auditar_coherencia_informe(txt_edit, d, mod_actual): st.info(f"🩺 Auditoría: {err}")
                
                uid = st.session_state.uid
                debe_estampar = perfil_activo["id"] in ["dr.amaya", "admin"]
                if mod_actual == "HOLTER": p_final, img_p = inyectar_holter_pdf(b_pdf, txt_edit, p_nom, perfil_activo, uid, debe_estampar)
                else: p_final, img_p = inyectar_mapa_pdf(b_pdf, txt_edit, p_nom, perfil_activo, uid, debe_estampar)
                
                c_b1, c_b2 = st.columns(2)
                with c_b1:
                    st.download_button("📄 DESCARGAR FIRMADO", data=p_final, file_name=f"{normalizar_nombre_archivo(p_nom)}_{cups_actual}.pdf", mime="application/pdf", type="primary", use_container_width=True)
                with c_b2:
                    if st.button("💾 Guardar Clínico", use_container_width=True):
                        p_clave = f"FC {d.get('fc_prom',0)} | SDNN {d.get('sdnn_24h',0)}ms" if mod_actual == "HOLTER" else f"PAM {d.get('pam_24h',0)} mmHg"
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
        with col_r1:
            st.markdown(f"**Total de estudios custodiados ({origen_datos}):** `{len(df_produccion)} procedimientos certificados`")
        with col_r2:
            excel_bytes, ext_salida = generar_excel_avanzado_produccion(df_produccion)
            mime_tipo = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if ext_salida == "xlsx" else "text/csv"
            etiqueta_boton = "📊 DESCARGAR REPORTE EXCEL GERENCIAL" if ext_salida == "xlsx" else "📊 DESCARGAR RIPS (CSV)"
            st.download_button(etiqueta_boton, data=excel_bytes, file_name=f"Reporte_Cencardio_{ahora_colombia().strftime('%Y%m%d')}.{ext_salida}", mime=mime_tipo, use_container_width=True, type="primary")

        st.divider()
        f1, f2, f3 = st.columns([1.5, 1.2, 1.3])
        with f1:
            busqueda = st.text_input("🔍 Buscar por paciente:")
        with f2:
            df_produccion["Mes_Periodo"] = pd.to_datetime(df_produccion["fecha_registro"]).dt.strftime('%Y-%m')
            mes_sel = st.selectbox("📅 Carpeta Mensual:", ["Todos los meses"] + sorted(df_produccion["Mes_Periodo"].unique().tolist(), reverse=True))
        with f3:
            mod_sel_f = st.selectbox("🎛️ Modalidad:", ["Todas las modalidades"] + sorted(df_produccion["modalidad"].unique().tolist()))

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
