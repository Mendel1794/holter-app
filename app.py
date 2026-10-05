import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(
    page_title="Lector Holter - CEN CARDIO",
    page_icon="🫀",
    layout="centered"
)

# ==========================================
# 1. USUARIOS Y CONTRASEÑAS AUTORIZADOS
# ==========================================
USUARIOS_AUTORIZADOS = {
    "dr.amaya": "Cardio2025*",
    "admin": "HolterClaveSegura123"
}

# ==========================================
# 2. CONTROL DE SESIÓN Y LOGIN
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = ""

def cerrar_sesion():
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""

if not st.session_state.autenticado:
    st.title("🫀 Acceso al Sistema de Lectura Holter")
    st.write("Ingresa tus credenciales autorizadas para continuar.")

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
                st.error("❌ Usuario o contraseña incorrectos. Verifica e intenta de nuevo.")

    st.stop()

# ==========================================
# 3. INTERFAZ Y PROCESAMIENTO
# ==========================================
with st.sidebar:
    st.write(f"👤 Conectado: **{st.session_state.usuario_actual}**")
    if st.button("Cerrar Sesión"):
        cerrar_sesion()
        st.rerun()
    st.divider()

    # Lee la API Key desde Secrets o desde un campo manual
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        api_key = st.text_input("Gemini API Key:", type="password", help="Pega aquí tu clave si no la pusiste en Secrets")

st.title("🫀 Interpretación Automatizada de Holter")
st.write("Sube el PDF emitido por el equipo Spacelabs para generar la lectura clínica individualizada.")

uploaded_file = st.file_uploader("Cargar estudio Holter (PDF)", type=["pdf"])

PROMPT_CARDIOLOGIA = """
Eres un médico cardiólogo experto. Analiza detalladamente TODO el documento PDF de este estudio Holter (incluyendo tablas de arritmias ventriculares y supraventriculares, frecuencias cardíacas, episodios de ST, intervalos QT/QTc, pausas y variabilidad SDNN).

Debes generar la lectura clínica individualizada con la estructura exacta de 10 puntos del Dr. William Amaya Ramirez, adaptando CADA PUNTO estrictamente a los hallazgos reales de ESTE paciente:

INTERPRETACIÓN TEST HOLTER

1. Ritmo de base y frecuencia cardiaca promedio (especificar ritmo sinusal o el ritmo base, frecuencia promedio en lpm y rango de FC mínima y máxima registradas).
2. Intervalos PR normal y QTc (indicar si son normales o si hay prolongación del QTc, indicando los valores en ms).
3. Alteraciones isquémicas del segmento ST (si hay infradesnivel o supradesnivel, reportar número de episodios, desviación máxima en mm y hora; si no hubo cambios, indicar 'Sin alteraciones isquémicas del segmento ST').
4. Conducción AV (indicar 'Sin Alteración en la conducción AV' o detallar bloqueos AV si existen).
5. Conducción intraventricular (indicar 'Sin Alteración en la conducción intraventricular' o detallar si hay bloqueo de rama).
6. Alteración de los impulsos por ectopias (detallar según los datos reales: si son supraventriculares y/o ventriculares, monomorfas/polimorfas, cantidad total o porcentaje, presencia de duplas, taquicardias o secuencias; si no hubo, indicar 'Sin alteración ectópica significativa').
7. Síntomas (indicar 'No refirió síntomas' o reportar los síntomas consignados en el diario del paciente).
8. Variabilidad de la FC (determinar si está 'conservada', 'disminuida' o 'severamente disminuida' según el SDNN de 24 horas).
9. Pausas (indicar 'Sin pausas significativas' o el conteo y duración de las pausas encontradas).
10. Riesgo del paciente SDNN a 24 HRS (clasificar según el valor del SDNN: Bajo riesgo >100ms, Riesgo medio 50-100ms, o Alto riesgo <50ms, indicando el número exacto en ms).

DR. WILLIAM AMAYA RAMIREZ
INTERNISTA - CARDIÓLOGO
RM 79.502.624 SDS

Nota: Sé exacto con los valores numéricos y diagnósticos del documento. No inventes datos que no figuren en las tablas.
"""

if uploaded_file is not None:
    if not api_key:
        st.warning("⚠️ Debes configurar la API Key de Gemini en la barra lateral izquierda o en Secrets.")
    else:
        if st.button("Generar Lectura del Estudio", type="primary"):
            with st.spinner("Analizando trazados y tablas del paciente..."):
                try:
                    client = genai.Client(api_key=api_key)
                    pdf_bytes = uploaded_file.getvalue()

                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[
                            types.Part.from_bytes(data=pdf_bytes, mime_type='application/pdf'),
                            PROMPT_CARDIOLOGIA
                        ]
                    )

                    st.success("✅ Lectura generada exitosamente.")
                    informe = st.text_area("Resultado (editable antes de copiar o imprimir):", value=response.text, height=380)

                    st.download_button(
                        label="Descargar Informe (.txt)",
                        data=informe,
                        file_name=f"Lectura_{uploaded_file.name}.txt",
                        mime="text/plain"
                    )
                except Exception as e:
                    st.error(f"Error al analizar el estudio: {e}")
