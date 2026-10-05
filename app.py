import streamlit as st
from pypdf import PdfReader
import fitz  # PyMuPDF
import re
import base64
import os
import io
import sqlite3
from datetime import datetime

st.set_page_config(
    page_title="Centro Cardiovascular Colombiano Cencardio",
    page_icon="🫀",
    layout="wide"
)

# ==========================================
# UTILIDAD: NORMALIZAR NOMBRE DE ARCHIVO
# ==========================================
def normalizar_nombre_archivo(nombre):
    limpio = re.sub(r'[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s]', ' ', nombre)
    limpio = re.sub(r'\s+', '_', limpio).strip('_')
    return limpio if limpio else "PACIENTE"

# ==========================================
# BASE DE DATOS LOCAL (HISTORIAL CLÍNICO)
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
            pdf_blob BLOB
        )
    """)
    conn.commit()
    conn.close()

init_db()

def guardar_estudio_db(nombre, fc, sdnn, medico, texto, pdf_bytes):
    conn = sqlite3.connect("historial_holter.db")
    c = conn.cursor()
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("""
        INSERT INTO estudios (fecha_registro, paciente_nombre, fc_prom, sdnn, medico_firmante, informe_texto, pdf_blob)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (fecha_actual, nombre, fc, sdnn, medico, texto, pdf_bytes))
    conn.commit()
    conn.close()

def obtener_historial_db():
    conn = sqlite3.connect("historial_holter.db")
    c = conn.cursor()
    c.execute("SELECT id, fecha_registro, paciente_nombre, fc_prom, sdnn, medico_firmante, pdf_blob FROM estudios ORDER BY id DESC")
    filas = c.fetchall()
    conn.close()
    return filas

def eliminar_estudio_db(estudio_id):
    conn = sqlite3.connect("historial_holter.db")
    c = conn.cursor()
    c.execute("DELETE FROM estudios WHERE id = ?", (estudio_id,))
    conn.commit()
    conn.close()

# ==========================================
# GESTIÓN DE FONDO PERSONALIZADO
# ==========================================
def cargar_fondo():
    for ext in ["fondo.jpg", "fondo.png", "fondo.jpeg"]:
        if os.path.exists(ext):
            with open(ext, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            mime = "png" if ext.endswith("png") else "jpeg"
            return f"""
            <style>
            .stApp {{
                background-image: linear-gradient(rgba(240, 248, 252, 0.90), rgba(240, 248, 252, 0.90)), 
                                  url("data:image/{mime};base64,{b64}");
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }}
            </style>
            """
    return """
    <style>
    .stApp { background: linear-gradient(135deg, #e8f4f8 0%, #f4f8fa 50%, #eef5f9 100%); }
    </style>
    """

st.markdown(cargar_fondo(), unsafe_allow_html=True)

# ==========================================
# LIMPIEZA DE INTERFAZ Y ESTILOS AVANZADOS
# ==========================================
st.markdown("""
    <style>
    /* 1. Eliminar franjas y recuadros superiores de Streamlit */
    header[data-testid="stHeader"] {
        background-color: transparent !important;
    }
    div[data-testid="stDecoration"] {
        display: none !important;
    }
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
    }

    /* 2. Eliminar el borde rígido que Streamlit le pone a los formularios */
    div[data-testid="stForm"] {
        border: none !important;
        padding: 0 !important;
    }

    /* 3. Tarjeta de Login limpia y profesional */
    .login-box {
        background: rgba(255, 255, 255, 0.95);
        border: 1px solid rgba(203, 213, 225, 0.8);
        border-radius: 16px;
        padding: 2.5rem;
        box-shadow: 0 10px 25px -5px rgba(12, 74, 110, 0.12);
        max-width: 580px;
        margin: 1rem auto;
    }

    /* 4. Tipografías y botones clínicos */
    h1 {
        color: #0c4a6e !important;
        font-weight: 800 !important;
        letter-spacing: -0.5px;
    }
    h2, h3 { color: #0369a1 !important; }
    [data-testid="stMetricValue"] { color: #0284c7 !important; font-weight: 700; }
    
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 1.2rem;
    }
    textarea {
        background-color: #ffffff !important;
        border: 1px solid #94a3b8 !important;
        border-radius: 8px !important;
        font-family: monospace !important;
        font-size: 13px !important;
    }
    .preview-container {
        border: 2px solid #cbd5e1;
        border-radius: 8px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        background: white;
        padding: 6px;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. PERFILES MÉDICOS Y SELECTOR DE ACCESO
# ==========================================
PERFILES_MEDICOS = {
    "dr.suarez": {
        "etiqueta": "👨‍⚕️ Dr. Martin Suárez Arámbula (Cardiólogo Hemodinamista)",
        "clave": "Suarez2026*",
        "nombre_completo": "DR. MARTIN SUÁREZ ARÁMBULA",
        "especialidad": "MÉDICO INTERNISTA - CARDIÓLOGO HEMODINAMISTA",
        "registro": "RM 13491094"
    },
    "dr.amaya": {
        "etiqueta": "👨‍⚕️ Dr. William Amaya Ramirez (Internista - Cardiólogo)",
        "clave": "Cardio2025*",
        "nombre_completo": "DR. WILLIAM AMAYA RAMIREZ",
        "especialidad": "INTERNISTA - CARDIÓLOGO",
        "registro": "RM 79.502.624 SDS"
    },
    "dra.cardio": {
        "etiqueta": "👩‍⚕️ Dra. Paola Figueroa (Cardióloga)",
        "clave": "Cardio2026*",
        "nombre_completo": "DRA. PAOLA FIGUEROA",
        "especialidad": "MÉDICO ESPECIALISTA EN CARDIOLOGÍA",
        "registro": "RM 52.890.123 SDS"
    },
    "admin": {
        "etiqueta": "⚙️ Administrador General del Sistema",
        "clave": "HolterClaveSegura123",
        "nombre_completo": "DR. MARTIN SUÁREZ ARÁMBULA",
        "especialidad": "MÉDICO INTERNISTA - CARDIÓLOGO HEMODINAMISTA",
        "registro": "RM 13491094"
    }
}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = ""

def cerrar_sesion():
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""

# PANTALLA DE ACCESO REFORMULADA Y CENTRADA
if not st.session_state.autenticado:
    col_v1, col_center, col_v2 = st.columns([1, 2.2, 1])
    
    with col_center:
        st.markdown("""
            <div class="login-box">
                <div style="text-align: center; margin-bottom: 1.5rem;">
                    <div style="font-size: 2.6rem; margin-bottom: 0.3rem;">🫀</div>
                    <div style="color: #0c4a6e; font-size: 1.35rem; font-weight: 800; line-height: 1.3; text-transform: uppercase;">
                        CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO
                    </div>
                    <div style="color: #0284c7; font-size: 0.95rem; font-weight: 600; margin-top: 0.4rem;">
                        Portal Médico de Interpretación y Lectura Holter
                    </div>
                </div>
        """, unsafe_allow_html=True)

        opciones_selector = {datos["etiqueta"]: usuario_key for usuario_key, datos in PERFILES_MEDICOS.items()}

        with st.form("form_login"):
            seleccion_etiqueta = st.selectbox(
                "Especialista Responsable:",
                options=list(opciones_selector.keys())
            )
            clave = st.text_input("Contraseña de Acceso:", type="password")
            boton_ingresar = st.form_submit_button("Ingresar al Portal", type="primary", use_container_width=True)

            if boton_ingresar:
                usuario_id = opciones_selector[seleccion_etiqueta]
                if PERFILES_MEDICOS[usuario_id]["clave"] == clave:
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = usuario_id
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta para el especialista seleccionado.")

        st.markdown('</div>', unsafe_allow_html=True)

    st.stop()

perfil_activo = PERFILES_MEDICOS[st.session_state.usuario_actual]

# ==========================================
# 2. MOTOR CLÍNICO SPACELABS
# ==========================================
with st.sidebar:
    st.write(f"👤 Especialista: **{perfil_activo['nombre_completo']}**")
    st.caption(f"{perfil_activo['especialidad']}\n{perfil_activo['registro']}")
    if st.button("Cerrar Sesión", use_container_width=True):
        cerrar_sesion()
        st.rerun()
    st.divider()
    st.caption("CENCARDIO - Sistema Clínico Integral v5.5")

st.markdown("""
    <div style="margin-bottom: 1rem;">
        <h1 style="margin: 0; font-size: 1.9rem;">🫀 CENTRO CARDIOVASCULAR COLOMBIANO CENCARDIO</h1>
        <p style="color: #0369a1; font-weight: 600; margin-top: 0.2rem; font-size: 1.05rem;">
            Sistema de Lectura Automatizada y Generación de Informes Holter Spacelabs
        </p>
    </div>
""", unsafe_allow_html=True)

tab_procesar, tab_historial = st.tabs(["📥 Procesar Nuevo Estudio", "📁 Archivo Clínico e Historial"])

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

{perfil['nombre_completo']}
{perfil['especialidad']}
{perfil['registro']}"""

    return informe

def inyectar_y_generar_preview(pdf_bytes, texto_informe):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]

    rects_h = pagina1.search_for("Hallazgos:")
    rects_f = pagina1.search_for("Firma del médico")
    if not rects_f:
        rects_f = pagina1.search_for("Firma del operador")

    if rects_h and rects_f:
        x0 = rects_h[0].x0 + 2
        y0 = rects_h[0].y1 + 4
        y1 = rects_f[0].y0 - 10
        x1 = pagina1.rect.width - 36
        rect_hallazgos = fitz.Rect(x0, y0, x1, y1)
    else:
        rect_hallazgos = fitz.Rect(35, 510, 565, 730)

    font_size = 8.2
    exito = False
    while font_size >= 4.5:
        rc = pagina1.insert_textbox(
            rect_hallazgos,
            texto_informe,
            fontsize=font_size,
            fontname="helv",
            color=(0, 0, 0),
            align=fitz.TEXT_ALIGN_LEFT
        )
        if rc >= 0:
            exito = True
            break
        font_size -= 0.3

    if not exito:
        y_cursor = rect_hallazgos.y0
        for linea in texto_informe.split("\n"):
            if y_cursor < rect_hallazgos.y1:
                pagina1.insert_text(
                    fitz.Point(rect_hallazgos.x0, y_cursor),
                    linea,
                    fontsize=6.5,
                    fontname="helv",
                    color=(0, 0, 0)
                )
                y_cursor += 9.5

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
            with st.spinner("Extrayendo datos de Spacelabs..."):
                st.session_state.datos_actuales = extraer_datos_spacelabs(bytes_originales, uploaded_file.name)
                st.session_state.texto_informe = redactar_interpretacion(st.session_state.datos_actuales, perfil_activo)
                st.session_state.archivo_actual = uploaded_file.name

        datos = st.session_state.datos_actuales

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("FC Promedio", f"{datos['fc_prom']} lpm", f"Mín {datos['fc_min']} | Máx {datos['fc_max']}")
        c2.metric("Ectopias Ventriculares", f"{datos['ev_total']} EV", f"TV: {datos['tv_episodios']}")
        c3.metric("Ectopias Supraventriculares", f"{datos['esv_total']} ESV", f"TSV: {datos['tsv_episodios']}")
        c4.metric("SDNN (24 Horas)", f"{datos['sdnn_24h']} ms", f"ST: {datos['st_episodios']} ep.")

        st.divider()

        col_edicion, col_preview = st.columns([1, 1], gap="large")

        with col_edicion:
            nombre_confirmado = st.text_input(
                "👤 Nombre del Paciente (editable para el archivo y descarga):",
                value=datos['paciente']
            )
            paciente_nombre_archivo = normalizar_nombre_archivo(nombre_confirmado)

            st.subheader("📝 Edición de la Interpretación")
            informe_para_grabar = st.text_area(
                "Edita el texto antes de generar el documento final:",
                value=st.session_state.texto_informe,
                height=350
            )

            pdf_final, img_preview = inyectar_y_generar_preview(bytes_originales, informe_para_grabar)

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
                        pdf_final
                    )
                    st.success(f"✅ Guardado con éxito: {nombre_confirmado}")

            st.write("")
            if st.button("🔄 Descartar / Limpiar Estudio Actual", use_container_width=True):
                if "datos_actuales" in st.session_state:
                    del st.session_state["datos_actuales"]
                if "archivo_actual" in st.session_state:
                    del st.session_state["archivo_actual"]
                st.rerun()

        with col_preview:
            st.subheader("👁️ Vista Previa Oficial (Página 1)")
            st.markdown('<div class="preview-container">', unsafe_allow_html=True)
            st.image(img_preview, caption=f"Página 1 - {nombre_confirmado}", use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# PESTAÑA 2: ARCHIVO CLÍNICO E HISTORIAL
# ==========================================
with tab_historial:
    st.subheader("📂 Registro de Estudios Procesados")
    historial = obtener_historial_db()

    if not historial:
        st.info("Aún no hay estudios archivados en el sistema.")
    else:
        busqueda = st.text_input("🔍 Buscar paciente por nombre o apellido:", "")
        
        for item in historial:
            est_id, fecha, pac_nom, fc, sdnn, med, pdf_data = item
            
            if busqueda.lower() in pac_nom.lower():
                nom_archivo_copia = normalizar_nombre_archivo(pac_nom)
                with st.expander(f"👤 {pac_nom} | 📅 {fecha} | 👨‍⚕️ {med}"):
                    c_det1, c_det2, c_desc, c_del = st.columns([2, 2, 2, 1.5])
                    c_det1.write(f"**FC Media:** {fc} lpm")
                    c_det2.write(f"**SDNN 24h:** {sdnn} ms")
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
