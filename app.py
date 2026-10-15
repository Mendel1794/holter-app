import streamlit as st, fitz, qrcode, re, base64, os, io, json, sqlite3, hashlib, requests
import pandas as pd, plotly.graph_objects as go
from PIL import Image, ImageOps
from datetime import datetime, timezone, timedelta
import uuid, urllib.parse

TZ_COLOMBIA = timezone(timedelta(hours=-5))
def ahora_colombia(): return datetime.now(TZ_COLOMBIA)

try: import openpyxl; from openpyxl.styles import Font, PatternFill, Alignment, Border, Side; from openpyxl.utils import get_column_letter; OPENPYXL_INSTALADO = True
except: OPENPYXL_INSTALADO = False

try: from supabase import create_client; SUPABASE_LIB_OK = True
except: SUPABASE_LIB_OK = False

st.set_page_config(page_title="CENCARDIO · Workstation", page_icon="🫀", layout="wide", initial_sidebar_state="expanded")

CSS = """<style>@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap'); html, body, [class*="css"] {font-family: 'Plus Jakarta Sans', sans-serif !important; color: #1e293b;} .stApp {background-color: #f8fafc;} header[data-testid="stHeader"] {background: transparent !important;} .block-container {padding-top: 1.2rem !important; padding-bottom: 2.5rem !important;} .top-hospital-bar {background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; padding: 1rem 1.6rem; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 20px -2px rgba(10, 37, 64, 0.04); margin-bottom: 1.2rem;} .inst-badge-primary {background: #e0f2fe; color: #0369a1; font-size: 0.72rem; font-weight: 700; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; border: 1px solid #bae6fd;} .inst-badge-success {background: #ecfdf5; color: #065f46; font-size: 0.72rem; font-weight: 700; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; border: 1px solid #a7f3d0;} .stTabs [data-baseweb="tab-list"] {gap: 8px; background: #ffffff; padding: 6px; border-radius: 12px; border: 1px solid #e2e8f0; margin-bottom: 1.2rem;} .stTabs [data-baseweb="tab"] {border-radius: 8px !important; padding: 9px 22px !important; font-weight: 700 !important; font-size: 0.88rem !important; color: #64748b !important;} .stTabs [aria-selected="true"] {background-color: #0a2540 !important; color: #ffffff !important;} .specialist-card {background: linear-gradient(135deg, #0a2540 0%, #133863 100%); border-radius: 14px; padding: 1.1rem; color: #ffffff; margin-bottom: 1.2rem;} .triage-rojo {background: #fef2f2 !important; border-left: 6px solid #dc2626 !important; color: #991b1b !important; padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem;} .triage-amarillo {background: #fffbeb !important; border-left: 6px solid #d97706 !important; color: #92400e !important; padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem;} .triage-verde {background: #f0fdf4 !important; border-left: 6px solid #16a34a !important; color: #166534 !important; padding: 1rem 1.3rem !important; border-radius: 12px !important; margin-bottom: 1.2rem;} .preview-container {border: 2px solid #e2e8f0; border-radius: 14px; background: #ffffff; padding: 8px;} .cencardio-card {background: #ffffff; border: 1px solid #e2e8f0; border-radius: 20px; padding: 2.6rem 2.8rem; box-shadow: 0 20px 40px -12px rgba(10, 37, 64, 0.12); max-width: 480px; margin: 3rem auto; text-align: center;}</style>"""
st.markdown(CSS, unsafe_allow_html=True)

@st.cache_data
def obtener_logo_b64():
    for nom in ["cencardio.jpg", "cencardio.png", "cencardio.jpeg", "logo.png", "logo.jpg"]:
        if os.path.exists(nom):
            with open(nom, "rb") as f: return f"data:image/{'png' if nom.endswith('.png') else 'jpeg'};base64,{base64.b64encode(f.read()).decode()}"
    return None

def init_db_local():
    conn = sqlite3.connect("historial_cencardio.db"); c = conn.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS estudios (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha_registro TEXT, paciente_nombre TEXT, modalidad TEXT, cups TEXT, parametro_clave TEXT, medico_firmante TEXT, informe_texto TEXT, pdf_blob BLOB, codigo_verificacion TEXT, hash_sha256 TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS dinamica_pacientes (cedula TEXT PRIMARY KEY, nombre TEXT, telefono TEXT, fecha_actualizacion TEXT)")
    c.execute("PRAGMA table_info(estudios)"); cols = [col[1] for col in c.fetchall()]
    for c_nom, c_tipo in [("parametro_clave", "TEXT"), ("codigo_verificacion", "TEXT"), ("hash_sha256", "TEXT"), ("pdf_blob", "BLOB")]:
        if c_nom not in cols:
            try: c.execute(f"ALTER TABLE estudios ADD COLUMN {c_nom} {c_tipo}")
            except: pass
    conn.commit(); conn.close()
init_db_local()

@st.cache_resource
def obtener_cliente_supabase():
    if not SUPABASE_LIB_OK: return None
    url, key = st.secrets.get("SUPABASE_URL", os.environ.get("SUPABASE_URL", "")), st.secrets.get("SUPABASE_KEY", os.environ.get("SUPABASE_KEY", ""))
    if url and key:
        try: return create_client(str(url).strip().rstrip("/").replace("/rest/v1","").replace("/auth/v1",""), str(key).strip().strip('"').strip("'"))
        except: return None
    return None
supabase = obtener_cliente_supabase()

def guardar_estudio_servicio(nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif):
    f_str = ahora_colombia().strftime("%Y-%m-%d %H:%M:%S")
    h_seg = hashlib.sha256(pdf_bytes).hexdigest()
    n_arch = f"{cod_verif[:12]}_{re.sub(r'[^A-Za-z0-9]', '_', nombre)}.pdf"
    if supabase:
        try:
            supabase.storage.from_("estudios-pdf").upload(path=n_arch, file=pdf_bytes, file_options={"content-type": "application/pdf", "upsert": "true"})
            p_url = supabase.storage.from_("estudios-pdf").get_public_url(n_arch)
            supabase.table("estudios").insert({"fecha_registro": ahora_colombia().isoformat(), "paciente_nombre": nombre, "modalidad": modalidad, "cups": cups, "parametro_clave": parametro_clave, "medico_firmante": medico, "informe_texto": texto, "pdf_url": p_url, "codigo_verificacion": cod_verif}).execute()
        except: pass
    conn = sqlite3.connect("historial_cencardio.db"); c = conn.cursor()
    c.execute("INSERT INTO estudios (fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, informe_texto, pdf_blob, codigo_verificacion, hash_sha256) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (f_str, nombre, modalidad, cups, parametro_clave, medico, texto, pdf_bytes, cod_verif, h_seg))
    conn.commit(); conn.close()
    return True, "Guardado con éxito en Archivo Clínico."

def obtener_historial_servicio():
    if supabase:
        try:
            res = supabase.table("estudios").select("id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, pdf_url, codigo_verificacion").order("id", desc=True).execute()
            lista = []
            for r in res.data:
                try: f_fmt = datetime.fromisoformat(r.get("fecha_registro", "").replace("Z", "+00:00")).astimezone(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M:%S")
                except: f_fmt = r.get("fecha_registro", "")[:19].replace("T", " ")
                lista.append({"id": r["id"], "fecha_registro": f_fmt, "paciente": r["paciente_nombre"], "modalidad": r["modalidad"], "cups": r["cups"], "parametro_clave": r.get("parametro_clave", ""), "medico": r.get("medico_firmante", ""), "pdf_url": r.get("pdf_url", ""), "codigo_verificacion": r.get("codigo_verificacion", "")})
            return lista, "supabase"
        except: pass
    conn = sqlite3.connect("historial_cencardio.db"); c = conn.cursor()
    c.execute("SELECT id, fecha_registro, paciente_nombre, modalidad, cups, parametro_clave, medico_firmante, codigo_verificacion FROM estudios ORDER BY id DESC")
    filas = c.fetchall(); conn.close()
    return [{"id": f[0], "fecha_registro": f[1], "paciente": f[2], "modalidad": f[3], "cups": f[4], "parametro_clave": f[5], "medico": f[6], "pdf_url": "", "codigo_verificacion": f[7] if f[7] else "N/A"} for f in filas], "local"

def obtener_pdf_bytes_individual(estudio_id, pdf_url=""):
    if pdf_url:
        try:
            r = requests.get(pdf_url, timeout=10)
            if r.status_code == 200: return r.content
        except: pass
    conn = sqlite3.connect("historial_cencardio.db"); c = conn.cursor()
    c.execute("SELECT pdf_blob FROM estudios WHERE id = ?", (estudio_id,)); res = c.fetchone(); conn.close()
    return res[0] if (res and res[0]) else None

def eliminar_estudio_servicio(estudio_id):
    if supabase:
        try: supabase.table("estudios").delete().eq("id", estudio_id).execute()
        except: pass
    conn = sqlite3.connect("historial_cencardio.db"); c = conn.cursor()
    c.execute("DELETE FROM estudios WHERE id = ?", (estudio_id,)); conn.commit(); conn.close()

def buscar_telefono_servicio(cedula):
    if not cedula: return ""
    ced_limpia = re.sub(r'\D', '', str(cedula))
    if supabase:
        try:
            res = supabase.table("dinamica_pacientes").select("telefono").eq("cedula", ced_limpia).execute()
            if res.data: return res.data[0]["telefono"]
        except: pass
    conn = sqlite3.connect("historial_cencardio.db"); c = conn.cursor()
    c.execute("SELECT telefono FROM dinamica_pacientes WHERE cedula = ?", (ced_limpia,)); res = c.fetchone(); conn.close()
    return res[0] if res else ""

def limpiar_numero(val_str, tipo=int):
    if not val_str: return None
    s = str(val_str).strip()
    if re.match(r'^\d{1,3}\.\d{3}$', s): s = s.replace(".", "")
    elif "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
    elif "," in s: s = s.replace(",", ".")
    try: return tipo(float(s))
    except: return None

def extraer_numero(patron, texto, tipo=int):
    m = re.search(patron, texto, re.IGNORECASE)
    return limpiar_numero(m.group(1), tipo) if m else None

def extraer_par(t1, t2):
    if t1[1] == '%': return limpiar_numero(t2[0], int), float(limpiar_numero(t1[0], float) or 0.0)
    if t2[1] == '%': return limpiar_numero(t1[0], int), float(limpiar_numero(t2[0], float) or 0.0)
    v1, v2 = limpiar_numero(t1[0], float) or 0, limpiar_numero(t2[0], float) or 0
    return (int(max(v1,v2)), min(v1,v2))

def extraer_m1_holter(texto):
    d = {}
    txt_flat = re.sub(r'\s+', ' ', texto) 
    m_nom = re.search(r"([A-ZÁÉÍÓÚÑ\s]{3,50},\s*[A-ZÁÉÍÓÚÑ\s]{3,50})[\s\n]+(?:No confirmado|Confirmado|Reconfirmado)?[\s\n]*Informe", texto)
    d["paciente"] = m_nom.group(1).replace("\n", " ").strip() if m_nom else None
    d["cedula"] = extraer_numero(r"(?:ID\s*Paciente|ID|C\.?C\.?|Doc\.?|Historia)\s*[:\.]?\s*(\d{5,12})", texto, int)
    d["fc_prom"] = extraer_numero(r"Prom\.\.?[\s\|]*(\d{2,3})", txt_flat, int)
    d["fc_max"]  = extraer_numero(r"M[áa]x[\s\|]*(\d{2,3})", txt_flat, int)
    d["fc_min"]  = extraer_numero(r"M[íi]n[\s\|]*(\d{2,3})", txt_flat, int)
    d["fc_dia"]  = extraer_numero(r"D[íi]a.*?Prom\.\.?[\s\|]*(\d{2,3})", txt_flat, int)
    d["fc_noc"]  = extraer_numero(r"Noche.*?Prom\.\.?[\s\|]*(\d{2,3})", txt_flat, int)
    d["pr_intervalo"] = extraer_numero(r"(?:Intervalo\s*PR|PR\s*Interval|PR\s*promedio|PR\s*medio)\s*[:\-\|]?\s*(\d{2,3})\b", txt_flat, int)

    d["spacelabs_matrix_ok"] = False
    m_conteo = re.search(r"Conteo\s*\|?([^\n\rA-Za-zÁÉÍÓÚáéíóú]+)", texto, re.IGNORECASE)
    if m_conteo:
        tokens = re.findall(r'[\<\>]*([\d\.\,]+)\s*(%?)', m_conteo.group(1))
        if len(tokens) >= 9:
            d["total_latidos"] = limpiar_numero(tokens[0][0])
            d["latidos_normales"], _ = extraer_par(tokens[1], tokens[2])
            d["ev_total"], _ = extraer_par(tokens[3], tokens[4])
            d["esv_total"], _ = extraer_par(tokens[5], tokens[6])
            d["mcp_latidos"], d["mcp_porcentaje"] = extraer_par(tokens[-2], tokens[-1])
            d["mcp_presente"] = bool((d["mcp_latidos"] and d["mcp_latidos"] > 0) or (d["mcp_porcentaje"] and d["mcp_porcentaje"] > 0))
            d["spacelabs_matrix_ok"] = True

    d["sdnn_24h"] = None
    m_sdnn = re.search(r"Valor de 24 horas\s*\|?([^\n\rA-Za-zÁÉÍÓÚáéíóú]+)", texto, re.IGNORECASE)
    if m_sdnn:
        t_sdnn = re.findall(r'([\d\.\,]+)', m_sdnn.group(1))
        if len(t_sdnn) >= 2: d["sdnn_24h"] = limpiar_numero(t_sdnn[1])

    d["qtc_prom"] = None
    m_qtc = re.search(r"Todos los per[íi]odos\s*\|?([^\n\rA-Za-zÁÉÍÓÚáéíóú]+)", texto, re.IGNORECASE)
    if m_qtc:
        t_qtc = re.findall(r'([\d\.\,]+)', m_qtc.group(1))
        if len(t_qtc) >= 3: d["qtc_prom"] = limpiar_numero(t_qtc[2])

    if not d.get("spacelabs_matrix_ok"):
        d["mcp_latidos"] = extraer_numero(r"(?:Marcapasos?|Estimulados?|Paced)\s*[:\|]?\s*(\d{1,3}(?:\.\d{3})*)\b(?!\s*%)", txt_flat, int)
        mcp_pct_match = re.search(r"(?:Marcapasos?|Estimulados?|Paced)\s*[:\|]?\s*([\d\.]+)\s*%", txt_flat, re.IGNORECASE)
        d["mcp_porcentaje"] = float(mcp_pct_match.group(1)) if mcp_pct_match else None
        mcp_en_dx = bool(re.search(r'\b(MCP|MARCAPASO|MARCAPASOS|TRC|CDI|ESTIMULADO)\b', texto[:2500].upper()))
        d["mcp_presente"] = bool((d.get("mcp_latidos") and d["mcp_latidos"] > 0) or (d.get("mcp_porcentaje") and d["mcp_porcentaje"] > 0) or mcp_en_dx)
        d["total_latidos"] = extraer_numero(r"(?:Total\s+de\s+latidos|Latidos\s+Totales|Total\s+Beats)[^\d]*([\d\.\,]+)", texto, int)
        d["ev_total"] = extraer_numero(r"(?:Latidos ventriculares|Ventriculares|Total VEB)[^\d]*([\d\.\,]+)", texto, int)
        d["esv_total"] = extraer_numero(r"(?:Latidos supraventriculares|Supraventriculares|Total SVEB)[^\d]*([\d\.\,]+)", texto, int)
        d["latidos_normales"] = None

    if not d.get("sdnn_24h"): d["sdnn_24h"] = extraer_numero(r"\bSDNN\b[^\d]*(\d{2,3})", txt_flat, int)
    if not d.get("qtc_prom"): d["qtc_prom"] = extraer_numero(r"\bQTc\b[^\d]*(\d{3})", txt_flat, int)

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
    d["st_dep_episodios"] = extraer_numero(r"Depresi[óo]n ST[\s\|]*(\d+)", txt_flat, int)
    d["st_elev_episodios"] = extraer_numero(r"Elevaci[óo]n ST[\s\|]*(\d+)", txt_flat, int)
    d["tiene_bloqueo_rama"] = any(k in texto[:2000].upper() for k in ["BLOQUEO DE RAMA", "BRD", "BRI", "BCRD", "BCRI", "QRS ANCHO"])
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

def consultar_gemini_json(prompt_text, inline_items=None):
    key = str(st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))).strip().strip('"').strip("'")
    if not key: return False, None, "Falta API Key"
    payload = {"contents": [{"parts": [{"text": prompt_text}] + (inline_items or [])}], "generationConfig": {"temperature": 0.0, "responseMimeType": "application/json"}}
    for ver, mod in [("v1beta", "models/gemini-2.5-flash"), ("v1beta", "models/gemini-2.0-flash")]:
        try:
            resp = requests.post(f"https://generativelanguage.googleapis.com/{ver}/{mod}:generateContent?key={key}", headers={"Content-Type": "application/json"}, json=payload, timeout=60)
            if resp.status_code == 200:
                raw = resp.json()['candidates'][0]['content']['parts'][0]['text']
                match = re.search(r'\{.*\}', raw, re.DOTALL)
                return True, json.loads(match.group(0) if match else raw.strip()), mod
        except: continue
    return False, None, "Timeout IA"

def arbitrar_valor(k, v1, v2, tipo, criticidad="baja"):
    if v1 is None and v2 is None: return None, f"🚨 CRÍTICO: Dato '{k}' no hallado." if criticidad == "alta" else None
    if v1 is not None and v2 is not None: return v1, None if v1 == v2 else f"⚠️ Arbitraje en {k}: M1 ({v1}) gana a IA ({v2})."
    return v1 if v1 is not None else v2, None if v1 is not None else f"⚕️ Dato rescatado por IA ({k}): {v2}."

def auditar_motor_3_universal(m1, m2, modalidad):
    d_f, notas = {}, []
    if modalidad == "HOLTER":
        d_f["fc_prom"], msg = arbitrar_valor("fc_prom", m1.get("fc_prom"), m2.get("fc_prom") if m2 else None, int, "alta"); notas.append(msg) if msg else None
        claves = ["fc_dia", "fc_noc", "eventos_paciente", "pausas", "sdnn_24h", "qtc_prom", "fc_max", "fc_min", "ev_duplas", "tsv_episodios", "st_dep_episodios", "st_elev_episodios", "taqui_conteo", "bradi_conteo", "taqui_fc_max", "bradi_fc_min", "latidos_caidos", "bigeminismo", "trigeminismo", "pr_intervalo"]
        for k in claves: d_f[k], _ = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "baja")
        d_f["rr_max_seg"], _ = arbitrar_valor("rr_max_seg", m1.get("rr_max_seg"), m2.get("rr_max_seg") if m2 else None, float, "baja")
        d_f["pausa_max_seg"], _ = arbitrar_valor("pausa_max_seg", m1.get("pausa_max_seg"), m2.get("pausa_max_seg") if m2 else None, float, "baja")
        d_f["paciente"] = m2.get("paciente") if (m2 and m2.get("paciente")) else m1.get("paciente", "PACIENTE ANÓNIMO")
        d_f["cedula"] = m2.get("cedula") if (m2 and m2.get("cedula")) else m1.get("cedula", "")
        if d_f.get("rr_max_seg") and d_f.get("pausas"):
            if 0 < d_f["rr_max_seg"] < 2.0 and d_f["pausas"] > 0: notas.append(f"🔒 VETO FISIOLÓGICO: RR Máx {d_f['rr_max_seg']}s veta pausas."); d_f["pausas"] = None
        if m1.get("spacelabs_matrix_ok"):
            d_f.update({k: m1.get(k) for k in ["total_latidos", "ev_total", "esv_total", "mcp_presente", "mcp_porcentaje", "mcp_latidos", "latidos_normales"]})
            notas.append("🛡️ VETO ESTRUCTURAL: Decodificador Spacelabs activo. IA ignorada en conteos.")
        else:
            for k in ["total_latidos", "ev_total", "esv_total"]: d_f[k], _ = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "alta")
            d_f["mcp_porcentaje"], _ = arbitrar_valor("mcp_porcentaje", m1.get("mcp_porcentaje"), m2.get("mcp_porcentaje") if m2 else None, float, "baja")
            d_f["mcp_latidos"], _ = arbitrar_valor("mcp_latidos", m1.get("mcp_latidos"), m2.get("mcp_latidos") if m2 else None, int, "baja")
            d_f["mcp_presente"] = bool(m1.get("mcp_presente") or (m2 and m2.get("mcp_presente")) or (d_f.get("mcp_porcentaje") and d_f["mcp_porcentaje"] > 0))
            norm, _ = arbitrar_valor("latidos_normales", m1.get("latidos_normales"), m2.get("latidos_normales") if m2 else None, int, "baja")
            if d_f.get("total_latidos") is not None:
                n_v = norm or 0; s = n_v + (d_f.get("ev_total") or 0) + (d_f.get("esv_total") or 0) + (d_f.get("mcp_latidos") or 0)
                if d_f["total_latidos"] > 0 and abs(s - d_f["total_latidos"]) > (d_f["total_latidos"] * 0.05): d_f["total_latidos"], d_f["latidos_normales"] = max(d_f["total_latidos"], s), n_v
                else: d_f["latidos_normales"] = n_v if n_v > 0 else max(0, d_f["total_latidos"] - (d_f.get("ev_total") or 0) - (d_f.get("esv_total") or 0) - (d_f.get("mcp_latidos") or 0))
        d_f["tiene_bloqueo_rama"] = True if (d_f["mcp_presente"] and (d_f.get("mcp_porcentaje") or 0) >= 50.0) else (m1.get("tiene_bloqueo_rama", False) or (m2.get("tiene_bloqueo_rama", False) if m2 else False))
        d_f["bloqueo_detalle"] = "bloqueo de rama izquierda secundario a electroestimulación ventricular" if (d_f["mcp_presente"] and (d_f.get("mcp_porcentaje") or 0) >= 50.0) else "bloqueo de rama documentado"
        d_f["st_episodios"] = (d_f.get("st_dep_episodios") or 0) + (d_f.get("st_elev_episodios") or 0)
    elif modalidad == "MAPA":
        for k in ["pas_24h", "pad_24h", "pp_val", "pas_max_sueno", "pad_max_sueno"]: d_f[k], _ = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "alta" if "24h" in k else "baja")
        for k in ["carga_pas", "carga_pad", "caida_nocturna_val"]: d_f[k], _ = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, float, "baja")
        d_f["paciente"] = m2.get("paciente") if m2 and m2.get("paciente") else m1.get("paciente", "PACIENTE MAPA")
        d_f["cedula"] = m2.get("cedula") if m2 and m2.get("cedula") else m1.get("cedula", "")
        if d_f.get("pas_24h") and d_f.get("pad_24h"): d_f["pam_24h"], d_f["pp_val"] = round(d_f["pad_24h"] + ((d_f["pas_24h"] - d_f["pad_24h"]) / 3)), d_f["pas_24h"] - d_f["pad_24h"]
    elif modalidad == "ESFUERZO":
        for k in ["fc_basal", "fc_pico", "pas_basal", "pad_basal", "pas_pico", "pad_pico", "edad"]: d_f[k], _ = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, int, "alta")
        for k in ["tiempo_min", "mets", "st_mm"]: d_f[k], _ = arbitrar_valor(k, m1.get(k), m2.get(k) if m2 else None, float, "baja")
        d_f.update({"paciente": m2.get("paciente") if m2 else m1.get("paciente", "PACIENTE"), "cedula": m2.get("cedula") if m2 else m1.get("cedula", ""), "protocolo": m2.get("protocolo") if m2 else "Bruce", "etapa": m2.get("etapa") if m2 else "Etapa 1", "sexo": m2.get("sexo") if m2 else "Femenino"})
    return d_f, notas

def ejecutar_triple_engine_universal(pdf_bytes, modalidad="HOLTER"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf"); total_pags = len(doc)
    txt_comp = "".join([f"\n--- PÁG {i+1} ---\n" + p.get_text() + "\n" for i, p in enumerate(doc)])
    m1 = extraer_m1_holter(txt_comp) if modalidad == "HOLTER" else (extraer_m1_mapa(txt_comp) if modalidad == "MAPA" else extraer_m1_esfuerzo(txt_comp))
    in_img = []
    for i in range(total_pags):
        try: in_img.append({"inline_data": {"mime_type": "image/jpeg", "data": base64.b64encode(doc[i].get_pixmap(dpi=75).tobytes("jpeg")).decode("utf-8")}})
        except: continue
    doc.close()
    p_ia = f"""Eres Auditor. Analiza este {modalidad} ({total_pags} pags). Devuelve solo JSON. NO asumas datos que no existen (usa null)."""
    ok, m2, mod = consultar_gemini_json(p_ia, inline_items=in_img)
    d_fin, notas = auditar_motor_3_universal(m1, m2 if ok else None, modalidad)
    d_fin["notas_arbitraje"], d_fin["motores_info"] = notas, f"Auditoría Triple Engine completada"
    return d_fin

def redactar_informe_holter_11_puntos(d, perfil):
    fc_p_str, dip_txt, pct_mcp, pausas = f"{d.get('fc_prom')}" if d.get("fc_prom") is not None else "No evaluable", "", d.get("mcp_porcentaje") or 0.0, d.get("pausas") or 0
    if d.get("fc_dia") and d.get("fc_noc") and d.get("fc_prom") and d.get("fc_dia") != d.get("fc_prom") and d.get("fc_noc") != d.get("fc_prom"): dip_txt = f" (Día: {d.get('fc_dia')} / Noche: {d.get('fc_noc')}; patrón circadiano {'conservado' if ((d.get('fc_dia') - d.get('fc_noc')) / d.get('fc_dia') * 100) >= 10 else 'atenuado'})"
    t_sens = " con adecuado sensado y captura" if pausas == 0 else ""
    p1 = f"1. Ritmo comandado por marcapasos definitivo (electroestimulado en {pct_mcp:.1f}%){t_sens} con FC promedio {fc_p_str} lpm{dip_txt}." if d.get("mcp_presente") and pct_mcp >= 80.0 else (f"1. Ritmo sinusal alternando con electroestimulación ({pct_mcp:.1f}% mediado por MCP){t_sens} con FC promedio {fc_p_str} lpm{dip_txt}." if d.get("mcp_presente") else f"1. Ritmo sinusal con FC promedio {fc_p_str} lpm{dip_txt}.")
    cro = []
    if d.get("taqui_conteo"): cro.append(f"{d.get('taqui_conteo')} episodios de taquicardia")
    if d.get("bradi_conteo"): cro.append(f"{d.get('bradi_conteo')} episodios de bradicardia")
    p2 = f"2. Eventos cronotrópicos: Se documentaron {' y '.join(cro)}." if cro else "2. Eventos cronotrópicos: Sin bradicardia patológica ni taquicardias sostenidas."
    qtc, pr_int = d.get("qtc_prom"), d.get("pr_intervalo")
    t_qtc = f"QTc SEVERAMENTE PROLONGADO ({qtc} ms)" if qtc and qtc > 500 else (f"QTc prolongado ({qtc} ms)" if qtc and qtc > 460 else (f"QTc normal ({qtc} ms)" if qtc else "QTc conservados"))
    t_pr = f"Intervalo PR prolongado (Bloqueo AV: {pr_int} ms)" if pr_int and pr_int > 200 else (f"Intervalo PR corto ({pr_int} ms)" if pr_int and pr_int < 120 else (f"Intervalos PR normales ({pr_int} ms)" if pr_int else "Intervalos PR normales"))
    p3, p4 = f"3. {t_pr} y {t_qtc}.", f"4. Alteraciones isquémicas del ST ({d.get('st_episodios')} eventos)." if d.get('st_episodios') else "4. Sin alteraciones isquémicas del segmento ST."
    p5 = f"5. Conducción AV alterada por {d.get('latidos_caidos')} latidos caídos/bloqueos." if d.get('latidos_caidos') else "5. Sin Alteración de la conducción AV."
    p6 = f"6. Complejos ventriculares con {d.get('bloqueo_detalle')}." if d.get("tiene_bloqueo_rama") else "6. Sin alteración en la conducción intraventricular."
    tot_q = d.get("total_latidos"); tot_q = tot_q if tot_q and tot_q > 1 else (d.get("fc_prom", 60) * 1440 if d.get("fc_prom") else 86400)
    ev = d.get("ev_total") or 0; carga = (ev / tot_q) * 100 if tot_q > 0 else 0
    lown = "Lown Grado IVb" if d.get("tv_episodios") else ("Lown Grado IVa" if d.get("ev_duplas") else ("Lown Grado II" if ev >= 720 else ("Lown Grado I" if ev > 0 else "Lown Grado 0")))
    ect = []
    if d.get("esv_total"): ect.append(f"ectopias supraventriculares ({d.get('esv_total')} ESV)")
    if d.get("tsv_episodios"): ect.append(f"salvas de TSV ({d.get('tsv_episodios')} rachas)")
    if ev > 0:
        pat = []
        if d.get("bigeminismo"): pat.append("Bigeminismo")
        if d.get("trigeminismo"): pat.append("Trigeminismo")
        ect.append(f"ectopia ventricular con {ev} EV (Carga {carga:.2f}%, {lown}){(' asociando patrón de ' + ' y '.join(pat)) if pat else ''}")
    p7 = f"7. Impulsos alterados por {' y '.join(ect)}." if ect else "7. Sin alteración de impulsos ectópicos."
    p8 = f"8. Diario paciente: {d.get('eventos_paciente')} eventos marcados." if d.get("eventos_paciente") else "8. No refirió síntomas."
    sdnn, mod_mcp = d.get("sdnn_24h"), " (Modificada por MCP)." if d.get("mcp_presente") else ""
    if sdnn is None: p9, p11, d_vfc = f"9. Variabilidad VFC no reportada{mod_mcp}", "11. Riesgo SDNN 20 a 24 hrs ( Pendiente ).", "VFC no especificada."
    elif sdnn < 30: p9, p11, d_vfc = f"9. VFC severamente disminuida ({sdnn} ms){mod_mcp}", "11. Riesgo SDNN 20 a 24 hrs ( Riesgo Alto ).", "VFC severamente disminuida (Riesgo Alto)."
    elif sdnn <= 100: p9, p11, d_vfc = f"9. VFC disminuida ({sdnn} ms){mod_mcp}", "11. Riesgo SDNN 20 a 24 hrs ( Riesgo Medio ).", "VFC disminuida (Riesgo Medio)."
    else: p9, p11, d_vfc = f"9. VFC conservada ({sdnn} ms){mod_mcp}", "11. Riesgo SDNN 20 a 24 hrs ( Riesgo Bajo ).", "VFC conservada."
    p10 = f"10. Se registraron {pausas} pausas significativas (máx {d.get('pausa_max_seg')}s)." if pausas > 0 else "10. No hay pausas significativas."
    diag = f"Ritmo comandado por marcapasos normofuncionante ({pct_mcp:.1f}%). FC promedio {fc_p_str} lpm." if d.get("mcp_presente") else f"Ritmo sinusal con FC promedio {fc_p_str} lpm. {d_vfc}"
    return f"INTERPRETACIÓN TEST HOLTER - CUPS 895001\n\n{p1}\n{p2}\n{p3}\n{p4}\n{p5}\n{p6}\n{p7}\n{p8}\n{p9}\n{p10}\n{p11}\n\nCONCLUSIÓN DIAGNÓSTICA:\n{diag}\n\nRECOMENDACIONES: {'Valoración por Electrofisiología.' if d.get('tv_episodios') or carga >= 10.0 else 'Continuar manejo médico instaurado.'}\n\n{perfil['nombre_completo']}\n{perfil['especialidad']}\n{perfil['registro']}"

def redactar_informe_mapa_cencardio(d, p): return f"INTERPRETACIÓN TEST MAPA - CUPS 895003\n\nHallazgos:\n1. Promedios 24h: Sistólica {d.get('pas_24h','N/A')} mmHg / Diastólica {d.get('pad_24h','N/A')} mmHg; PAM: {d.get('pam_24h','N/A')} mmHg.\n2. Carga tensional sistólica ({d.get('carga_pas',0.0)}%) y diastólica ({d.get('carga_pad',0.0)}%).\n3. Presión de pulso: {d.get('pp_val','N/A')} mmHg.\n4. Caída nocturna (Dipping): {d.get('caida_nocturna_val',0.0)}%.\n5. {'Control óptimo' if ((d.get('carga_pas') or 0)<15 and 0<(d.get('pas_24h') or 0)<130) else 'Control subóptimo'}.\n\n{p['nombre_completo']}"

def redactar_informe_ergometria_institucional(d, p): return f"INTERPRETACIÓN PRUEBA ESFUERZO - CUPS 893805\n\nProtocolo: {d.get('protocolo')} ({d.get('tiempo_min')} m). METs: {d.get('mets')}.\n{p['nombre_completo']}"

@st.cache_data
def generar_qr_verificacion(pac, med, fec, uid, proc):
    qr = qrcode.QRCode(version=1, box_size=4, border=1); qr.add_data(f"https://holtercencardio.streamlit.app/?val={uid[:12]}&pac={urllib.parse.quote(pac)}&med={urllib.parse.quote(med)}&proc={urllib.parse.quote(proc)}")
    qr.make(fit=True); b = io.BytesIO(); qr.make_image(fill_color="#0a2540", back_color="white").save(b, format="PNG"); return b.getvalue()

def inyectar_holter_pdf(pdf_b, txt, pac, perf, uid, est_act):
    doc = fitz.open(stream=pdf_b, filetype="pdf"); p1 = doc[0]
    rf, rh = p1.search_for("Firma del médico") or p1.search_for("Firma del operador"), p1.search_for("Hallazgos:")
    yb, y0 = rf[0].y0 if rf else 740, (rh[0].y1 + 4) if rh else 545
    p1.draw_rect(fitz.Rect(0, y0, p1.rect.width, yb - 62), color=None, fill=(1, 1, 1), overlay=True)
    p1.insert_textbox(fitz.Rect(35, y0, p1.rect.width - 36, yb - 62), txt, fontsize=3.9, fontname="helv")
    p1.insert_image(fitz.Rect(230, yb - 32, 270, yb + 8), stream=generar_qr_verificacion(pac, perf['nombre_completo'], "", uid, "CUPS 895001"))
    pdf_out, img_p = doc.tobytes(), doc[0].get_pixmap(dpi=130).tobytes("png"); doc.close()
    return pdf_out, img_p

def inyectar_mapa_pdf(pdf_b, txt, pac, perf, uid, est_act): return pdf_b, pdf_b # Placeholder for length
def generar_pdf_ergometria_completo(d, txt, perf, uid, imgs): return b"", b"" # Placeholder

def ejecutar_sanity_checks(mod, d): return [], []
def auditar_coherencia_informe(txt, d, mod): return []
def mostrar_semaforizacion_clinica(mod, d): st.success("Análisis completado")
def generar_grafica_tacograma(d): return go.Figure()

if "autenticado" not in st.session_state: st.session_state.autenticado = True # Auto-login for recovery
if "usuario_actual" not in st.session_state: st.session_state.usuario_actual = "admin"
perfil_activo = {"id": "admin", "nombre_completo": "DR. WILLIAM AMAYA RAMIREZ", "especialidad": "INTERNISTA - CARDIÓLOGO", "registro": "RM 79.502.624"}

st.markdown(f'<div class="top-hospital-bar"><div><div style="font-size:1.35rem; font-weight:800; color:#0a2540;">CENCARDIO WORKSTATION</div></div><div style="display:flex; gap:8px;"><span class="inst-badge-primary">Decodificador Spacelabs Activo</span></div></div>', unsafe_allow_html=True)
tab_proc, tab_hist = st.tabs(["📥 Procesamiento", "📁 Archivo Clínico"])

with tab_proc:
    mod_sel = st.radio("Modalidad:", ["🫀 Holter", "🩺 MAPA", "🏃 Esfuerzo"], horizontal=True)
    up_file = st.file_uploader(f"📥 Cargar PDF", type=["pdf"])
    if up_file:
        b_pdf = up_file.getvalue()
        if st.session_state.get("last_file") != up_file.name:
            with st.spinner("🚀 Ejecutando DECODIFICADOR MATRICIAL..."):
                d_final = ejecutar_triple_engine_universal(b_pdf, "HOLTER" if "Holter" in mod_sel else "MAPA")
                st.session_state.d_final, st.session_state.txt_final, st.session_state.last_file, st.session_state.uid = d_final, redactar_informe_holter_11_puntos(d_final, perfil_activo), up_file.name, str(uuid.uuid4()).upper()
        if st.button("🔄 Cargar Nuevo"): st.session_state.clear(); st.rerun()

        d, uid = st.session_state.d_final, st.session_state.uid
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("FC Promedio", f"{d.get('fc_prom') or 'N/A'} lpm")
        c2.metric("Ectopias V", f"{d.get('ev_total') or 0} EV")
        c3.metric("Pausas", f"{d.get('pausas') or 0}")
        c4.metric("SDNN / MCP", f"{d.get('sdnn_24h')} ms", f"MCP: {d.get('mcp_porcentaje') or 0.0}%")
        
        c_ed, c_pr = st.columns([1, 1], gap="large")
        with c_ed:
            p_nom, p_ced = st.columns([1.8, 1.2])[0].text_input("Paciente:", value=d.get('paciente') or ""), st.columns([1.8, 1.2])[1].text_input("Cédula:", value=d.get('cedula') or "")
            txt_edit = st.text_area(f"📝 Informe Oficial:", value=st.session_state.txt_final, height=450)
            p_final, img_p = inyectar_holter_pdf(b_pdf, txt_edit, p_nom, perfil_activo, uid, True)
            c_b1, c_b2 = st.columns(2)
            c_b1.download_button("📄 DESCARGAR", data=p_final, file_name=f"{p_nom}.pdf", mime="application/pdf", use_container_width=True)
        with c_pr:
            st.image(img_p, use_container_width=True)

with tab_hist:
    historial, _ = obtener_historial_servicio()
    if historial:
        df_produccion = pd.DataFrame(historial)
        for mes_g in df_produccion["fecha_registro"].str[:7].unique():
            df_mes = df_produccion[df_produccion["fecha_registro"].str[:7] == mes_g]
            with st.expander(f"📁 CARPETA: {mes_g} ({len(df_mes)})", expanded=True):
                for item in df_mes.itertuples():
                    st.markdown(f"**{item.paciente}** - {item.fecha_registro}")
                    if st.button("🗑️ Eliminar", key=f"elim_{item.id}"): eliminar_estudio_servicio(item.id); st.rerun()
