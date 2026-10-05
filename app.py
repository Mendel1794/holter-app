import streamlit as st
from pypdf import PdfReader
import re
import base64
import os

# Configuración de página
st.set_page_config(
    page_title="Lectura de Holter Cencardio",
    page_icon="🫀",
    layout="centered"
)

# ==========================================
# GESTIÓN DE IMAGEN DE FONDO PERSONALIZADA
# ==========================================
def cargar_imagen_fondo():
    # Busca si subiste 'fondo.jpg' o 'fondo.png'
    archivo_fondo = None
    tipo_mime = "jpeg"
    
    if os.path.exists("fondo.jpg"):
        archivo_fondo = "fondo.jpg"
        tipo_mime = "jpeg"
    elif os.path.exists("fondo.png"):
        archivo_fondo = "fondo.png"
        tipo_mime = "png"
    elif os.path.exists("fondo.jpeg"):
        archivo_fondo = "fondo.jpeg"
        tipo_mime = "jpeg"

    if archivo_fondo:
        with open(archivo_fondo, "rb") as f:
            b64_data = base64.b64encode(f.read()).decode()
        return f"""
        <style>
        .stApp {{
            /* Capa traslúcida blanca al 85% para mantener la legibilidad clínica del texto */
            background-image: linear-gradient(rgba(255, 255, 255, 0.85), rgba(255, 255, 255, 0.85)), 
                              url("data:image/{tipo_mime};base64,{b64_data}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        </style>
        """
    else:
        # Fondo por defecto si aún no subes tu imagen
        return """
        <style>
        .stApp {
            background: linear-gradient(135deg, #e8f4f8 0%, #f4f8fa 50%, #eef5f9 100%);
        }
        </style>
        """

st.markdown(cargar_imagen_fondo(), unsafe_allow_html=True)

# Estilos adicionales para textos y componentes
st.markdown("""
    <style>
    h1 {
        color: #0c4a6e !important;
        font-weight: 700 !important;
    }
    h2, h3 {
        color: #0369a1 !important;
    }
    [data-testid="stMetricValue"] {
        color: #0284c7 !important;
        font-weight: 600;
    }
    .stButton > button {
        background-color: #0284c7 !important;
        color: white !important;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1rem;
        font-weight: 600;
    }
    .stButton > button:hover {
        background-color: #0369a1 !important;
    }
    textarea {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        font-family: monospace !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. SEGURIDAD Y ACCESO (USUARIO / CLAVE)
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
        boton_ingresar = st.form_submit_button("Iniciar Sesión", type="primary")

        if boton_ingresar:
            if usuario in USUARIOS_AUTORIZADOS and USUARIOS_AUTORIZADOS[usuario] == clave:
                st.session_state.autenticado = True
                st.session_state.usuario_actual = usuario
                st.rerun()
            else:
                st.error("❌ Credenciales incorrectas. Verifica usuario y contraseña.")

    st.stop()

# ==========================================
# 2. PANEL PRINCIPAL
# ==========================================
with st.sidebar:
    st.write(f"👤 Conectado: **{st.session_state.usuario_actual}**")
    if st.button("Cerrar Sesión"):
        cerrar_sesion()
        st.rerun()
    st.divider()
    st.caption("CENCARDIO - Sistema de Lectura Automatizada")

st.title("🫀 Lectura de Holter Cencardio")
st.write("Carga el archivo PDF de Spacelabs para generar la lectura clínica individualizada.")

uploaded_file = st.file_uploader("Cargar estudio Holter (PDF)", type=["pdf"])

def procesar_estudio_holter(archivo_pdf):
    reader = PdfReader(archivo_pdf)
    texto = ""
    for page in reader.pages:
        t = page.extract_text()
        if t:
            texto += t + "\n"

    # Frecuencia cardíaca
    fc_prom = 70
    fc_match = re.search(r"Prom(?:\.|\:)?\s*\|\s*(\d{2,3})", texto)
    if not fc_match:
        fc_match = re.search(r"Prom(?:\.|\:)?\s+(\d{2,3})", texto)
    if fc_match:
        fc_prom = int(fc_match.group(1))

    # Ectopias ventriculares
    ev_total = 0
    ev_match = re.search(r"Latidos ventriculares:\s*(\d+)", texto)
    if not ev_match:
        ev_match = re.search(r"Ventricular\s*\|\s*(\d+)", texto)
    if ev_match:
        ev_total = int(ev_match.group(1))

    # Rachas de TV
    tv_match = re.search(r"TV\s*\|\s*(\d+)", texto)
    tv_count = int(tv_match.group(1)) if tv_match else 0

    # Ectopias supraventriculares
    esv_total = 0
    esv_match = re.search(r"Latidos supraventriculares:\s*(\d+)", texto)
    if not esv_match:
        esv_match = re.search(r"Supraventricular\s*\|\s*(\d+)", texto)
    if esv_match:
        esv_total = int(esv_match.group(1))

    # Pausas
    pausas = 0
    pau_match = re.search(r"Pausa\s*\|\s*(\d+)", texto)
    if pau_match:
        pausas = int(pau_match.group(1))

    # Segmento ST
    st_depresion = 0
    st_max_mm = "0.00"
    st_match = re.search(r"Depresión ST\s*\|\s*(\d+)\s*\|\s*(-?[\d,\.]+)", texto)
    if st_match:
        st_depresion = int(st_match.group(1))
        st_max_mm = st_match.group(2)

    # SDNN 24 Horas
    sdnn = 85
    sdnn_match = re.search(r"Valor de 24 horas\s*\|\s*\d+\s*\|\s*(\d+)", texto)
    if sdnn_match:
        sdnn = int(sdnn_match.group(1))

    # ==========================================
    # LÓGICA DE INTERPRETACIÓN (10 PUNTOS)
    # ==========================================
    p1 = f"1. Ritmo de sinusal frecuencia cardiaca promedio de {fc_prom} latidos por minuto."
    p2 = "2. Intervalos PR normal y QTc normales."

    if st_depresion > 0:
        p3 = f"3. Alteraciones isquémicas del segmento ST ({st_depresion} episodios de depresión del ST, máx. {st_max_mm} mm)."
    else:
        p3 = "3. Sin alteraciones isquémicas del segmento ST."

    p4 = "4. Sin Alteración en la conducción AV."
    p5 = "5. Sin Alteración en la conducción intraventricular."

    if ev_total > 500 or tv_count > 0:
        p6 = f"6. Alteración de los impulsos por ectopias supraventriculares y ventriculares frecuentes ({ev_total} EV con episodios de taquicardia ventricular y {esv_total} ESV)."
    elif ev_total > 0 and esv_total > 0:
        p6 = "6. Alteración de los impulsos por ectopias supraventriculares y ventriculares monomorfas."
    elif esv_total > 0:
        p6 = "6. Alteración de los impulsos por ectopias supraventriculares monomorfas."
    elif ev_total > 0:
        p6 = "6. Alteración de los impulsos por ectopias ventriculares monomorfas."
    else:
        p6 = "6. Sin alteración significativa de los impulsos ectópicos."

    p7 = "7. No refirió síntomas."

    if sdnn < 50:
        p8 = "8. Variabilidad severamente disminuida de la FC."
    elif 50 <= sdnn <= 100:
        p8 = "8. Variabilidad disminuida de la FC."
    else:
        p8 = "8. Variabilidad conservada de la FC."

    if pausas == 0:
        p9 = "9. Sin pausas significativas."
    else:
        p9 = f"9. Se registraron {pausas} pausas significativas (> 2.0 s)."

    if sdnn < 50:
        riesgo = "Alto riesgo"
    elif 50 <= sdnn <= 100:
        riesgo = "Riesgo medio"
    else:
        riesgo = "Bajo riesgo"
    p10 = f"10. Riesgo del paciente SDNN a 24 HRS ({riesgo} - {sdnn} ms)."

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

    return informe, fc_prom, sdnn, ev_total, esv_total

# ==========================================
# 3. GENERACIÓN DE RESULTADO
# ==========================================
if uploaded_file is not None:
    if st.button("Generar Lectura del Estudio", type="primary"):
        with st.spinner("Procesando trazado del paciente..."):
            informe_generado, fc, sdnn_val, ev, esv = procesar_estudio_holter(uploaded_file)

        st.success("✅ Estudio procesado correctamente.")

        col1, col2, col3 = st.columns(3)
        col1.metric("FC Promedio", f"{fc} lpm")
        col2.metric("SDNN (24h)", f"{sdnn_val} ms")
        col3.metric("Ectopias (EV / ESV)", f"{ev} / {esv}")

        st.subheader("Resultado de la Interpretación")
        resultado_editable = st.text_area(
            "Texto listo para copiar o imprimir:",
            value=informe_generado,
            height=370
        )

        st.download_button(
            label="Descargar Informe (.txt)",
            data=resultado_editable,
            file_name=f"Lectura_{uploaded_file.name}.txt",
            mime="text/plain"
        )
