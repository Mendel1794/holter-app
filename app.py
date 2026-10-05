import streamlit as st
from pypdf import PdfReader
import fitz  # PyMuPDF
import re
import base64
import os
import io

st.set_page_config(
    page_title="Lectura de Holter Cencardio",
    page_icon="🫀",
    layout="wide"
)

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
                background-image: linear-gradient(rgba(255, 255, 255, 0.90), rgba(255, 255, 255, 0.90)), 
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

st.markdown("""
    <style>
    h1 { color: #0c4a6e !important; font-weight: 700 !important; }
    h2, h3 { color: #0369a1 !important; }
    [data-testid="stMetricValue"] { color: #0284c7 !important; font-weight: 700; }
    .stButton > button {
        background-color: #0284c7 !important;
        color: white !important;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 1.5rem;
    }
    textarea {
        background-color: #ffffff !important;
        border: 1px solid #94a3b8 !important;
        border-radius: 8px !important;
        font-family: monospace !important;
        font-size: 13px !important;
    }
    /* Marco elegante para la vista previa del documento */
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
# 1. CONTROL DE ACCESO
# ==========================================
USUARIOS_AUTORIZADOS = {
    "dr.amaya": "Cardio2025*",
    "admin": "HolterClaveSegura123"
}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = ""

def cerrar_sesion():
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""

if not st.session_state.autenticado:
    st.title("🫀 Lectura de Holter Cencardio")
    st.write("Ingresa tus credenciales autorizadas para acceder al portal.")

    with st.form("form_login"):
        usuario = st.text_input("Usuario")
        clave = st.text_input("Contraseña", type="password")
        if st.form_submit_button("Iniciar Sesión", type="primary"):
            if usuario in USUARIOS_AUTORIZADOS and USUARIOS_AUTORIZADOS[usuario] == clave:
                st.session_state.autenticado = True
                st.session_state.usuario_actual = usuario
                st.rerun()
            else:
                st.error("❌ Credenciales incorrectas. Verifica usuario y contraseña.")
    st.stop()

# ==========================================
# 2. MOTOR CLÍNICO SPACELABS
# ==========================================
with st.sidebar:
    st.write(f"👤 Conectado: **{st.session_state.usuario_actual}**")
    if st.button("Cerrar Sesión"):
        cerrar_sesion()
        st.rerun()
    st.divider()
    st.caption("CENCARDIO - Sistema Holter Profesional v4.0")

st.title("🫀 Lectura de Holter Cencardio")
st.write("Carga el archivo PDF de Spacelabs. El sistema extraerá los datos, redactará la lectura y mostrará la vista previa del documento diligenciado.")

uploaded_file = st.file_uploader("Cargar estudio Holter (PDF)", type=["pdf"])

def limpiar_numero(val_str):
    if not val_str:
        return 0
    val_str = val_str.replace(".", "").replace(",", ".")
    try:
        return int(float(val_str))
    except:
        return 0

def extraer_datos_spacelabs(pdf_bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    texto = ""
    for page in reader.pages:
        t = page.extract_text()
        if t:
            texto += t + "\n"

    datos = {}

    paciente_match = re.search(r"ID paciente:.*?\n([A-ZÁÉÍÓÚÑ\s,]+)\nInforme Holter", texto)
    if paciente_match:
        datos["paciente"] = paciente_match.group(1).replace("\n", " ").strip()
    else:
        datos["paciente"] = "Paciente_Estudio"

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

def redactar_interpretacion(d):
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

DR. WILLIAM AMAYA RAMIREZ
INTERNISTA - CARDIÓLOGO
RM 79.502.624 SDS"""

    return informe

# ==========================================
# 3. INYECCIÓN Y PREVISUALIZADOR
# ==========================================
def inyectar_y_generar_preview(pdf_bytes, texto_informe):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina1 = doc[0]

    # Localización dinámica del recuadro
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

    # Auto-escalado de fuente
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

    # Renderizar la Página 1 como imagen PNG de alta definición (DPI 150)
    pix = pagina1.get_pixmap(dpi=150)
    img_preview = pix.tobytes("png")

    # Documento PDF completo modificado
    pdf_final_bytes = doc.tobytes()
    doc.close()

    return pdf_final_bytes, img_preview

# ==========================================
# 4. INTERFAZ DE USUARIO EN DOS COLUMNAS
# ==========================================
if uploaded_file is not None:
    bytes_originales = uploaded_file.getvalue()

    if "datos_actuales" not in st.session_state or st.session_state.get("archivo_actual") != uploaded_file.name:
        with st.spinner("Extrayendo datos de Spacelabs..."):
            st.session_state.datos_actuales = extraer_datos_spacelabs(bytes_originales)
            st.session_state.texto_informe = redactar_interpretacion(st.session_state.datos_actuales)
            st.session_state.archivo_actual = uploaded_file.name

    datos = st.session_state.datos_actuales

    # Métricas de verificación clínica
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("FC Promedio", f"{datos['fc_prom']} lpm", f"Mín {datos['fc_min']} | Máx {datos['fc_max']}")
    c2.metric("Ectopias Ventriculares", f"{datos['ev_total']} EV", f"TV: {datos['tv_episodios']}")
    c3.metric("Ectopias Supraventriculares", f"{datos['esv_total']} ESV", f"TSV: {datos['tsv_episodios']}")
    c4.metric("SDNN (24 Horas)", f"{datos['sdnn_24h']} ms", f"ST: {datos['st_episodios']} ep.")

    st.divider()

    # Distribución en 2 columnas: Edición a la izquierda, Vista previa a la derecha
    col_edicion, col_preview = st.columns([1, 1], gap="large")

    with col_edicion:
        st.subheader("📝 Edición de la Interpretación")
        informe_para_grabar = st.text_area(
            "Edita aquí el texto si requieres agregar comentarios (la vista previa se actualizará):",
            value=st.session_state.texto_informe,
            height=400
        )

        # Generar PDF y la imagen de vista previa
        pdf_final, img_preview = inyectar_y_generar_preview(bytes_originales, informe_para_grabar)

        st.download_button(
            label="📄 DESCARGAR PDF OFICIAL DILIGENCIADO",
            data=pdf_final,
            file_name=f"Holter_{datos['paciente'].replace(' ', '_')}_Firmado.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )

    with col_preview:
        st.subheader("👁️ Vista Previa en Vivo (Página 1)")
        st.markdown('<div class="preview-container">', unsafe_allow_html=True)
        st.image(
            img_preview, 
            caption=f"Página 1 - {datos['paciente']}", 
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
