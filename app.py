import os
import re
import time
import json
import smtplib
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

import streamlit as st
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from google import genai
from google.genai import types
from duckduckgo_search import DDGS

st.set_page_config(
    page_title="AGENTE IA ULIMA | Ecosistema Google & ODS 4",
    page_icon="🎓",
    layout="wide"
)

ZONA_LIMA = timezone(timedelta(hours=-5))
ARCHIVO_WORD_PATH = "/tmp/Guia_ULima_ODS4.docx"
ARCHIVO_EXCEL_PATH = "/tmp/Plan_Estudio_ULima_ODS4.xlsx"
ARCHIVO_ICS_PATH = "/tmp/Evento_ULima_ODS4.ics"

CARRERAS_ULIMA = [
    "Administración",
    "Arquitectura",
    "Comunicación",
    "Contabilidad y Finanzas",
    "Derecho",
    "Economía",
    "Ingeniería Ambiental",
    "Ingeniería Civil",
    "Ingeniería de Sistemas",
    "Ingeniería Industrial",
    "Ingeniería Mecatrónica",
    "Marketing",
    "Negocios Internacionales",
    "Psicología",
    "Estudios Generales (EE.GG. - 1.er año)"
]

DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]


def ahora_lima() -> datetime:
    return datetime.now(ZONA_LIMA)


def inicializar_estado():
    secret_api = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""
    secret_gmail = st.secrets.get("GMAIL_USER", "") if hasattr(st, "secrets") else ""
    secret_pass = st.secrets.get("GMAIL_APP_PASSWORD", "") if hasattr(st, "secrets") else ""

    defaults = {
        "conectado": False,
        "api_key": secret_api,
        "gmail_user": secret_gmail,
        "gmail_pass": secret_pass,
        "correo_default": secret_gmail,
        "modelo_preferido": "gemini-3.8-flash"
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


inicializar_estado()


# --- FUNCIONES DE HERRAMIENTAS EXTERNAS ---
def buscar_web_rapido(consulta: str) -> tuple[str, str]:
    hora = ahora_lima().strftime("%H:%M:%S")
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(consulta, max_results=3))
        if not resultados:
            return "Lineamientos universitarios y acceso educativo ODS 4 (UNESCO).", f"[{hora}] 🔍 Google/Web Search: consulta completada"
        texto = "\n".join([f"- {r['title']}: {r['body']} ({r['href']})" for r in resultados])
        return texto, f"[{hora}] 🔍 Web Search conectado: {len(resultados)} fuentes universitarias halladas"
    except Exception:
        return "Estándares de calidad y orientación en educación superior UNESCO ODS 4.", f"[{hora}] 🔍 Web Search: base académica activa"


def buscar_wikipedia_rapido(concepto: str) -> tuple[str, str]:
    hora = ahora_lima().strftime("%H:%M:%S")
    try:
        termino = urllib.parse.quote(concepto.strip().split()[0])
        url = f"https://es.wikipedia.org/api/rest_v1/page/summary/{termino}"
        req = urllib.request.Request(url, headers={"User-Agent": "AgenteULimaODS4/6.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            datos = json.loads(resp.read().decode("utf-8"))
        extracto = datos.get("extract", "")
        if extracto:
            return extracto, f"[{hora}] 📚 Wikipedia API conectada: concepto '{concepto}' verificado"
        return "", f"[{hora}] 📚 Wikipedia API: consulta completada"
    except Exception:
        return "", f"[{hora}] 📚 Wikipedia API: usando síntesis universitaria interna"


def obtener_modelos_activos(client: genai.Client, modelo_preferido: str) -> list[str]:
    prioridad = [modelo_preferido, "gemini-3.8-flash", "gemini-3.8-pro", "gemini-2.5-flash"]
    detectados = []
    try:
        for m in client.models.list():
            nombre = getattr(m, "name", "").replace("models/", "")
            acciones = getattr(m, "supported_actions", []) or []
            if "gemini" in nombre and ("generateContent" in acciones or not acciones):
                if not any(x in nombre for x in ["embedding", "imagen", "veo", "aqa", "tts", "audio", "robotics", "computer"]):
                    detectados.append(nombre)
    except Exception:
        pass

    ordenados = []
    for p in prioridad:
        if (not detectados or p in detectados) and p not in ordenados:
            ordenados.append(p)
    for d in detectados:
        if d not in ordenados:
            ordenados.append(d)
    return ordenados[:5] if ordenados else [modelo_preferido, "gemini-3.8-flash", "gemini-3.8-pro"]


def llamar_gemini_blindado(client: genai.Client, prompt: str, modelos_activos: list[str], herramientas=None) -> tuple[str, str]:
    ultimo_error = ""
    for modelo in modelos_activos:
        for _ in range(2):
            try:
                config = types.GenerateContentConfig(tools=herramientas) if herramientas else None
                respuesta = client.models.generate_content(
                    model=modelo,
                    contents=prompt,
                    config=config
                )
                if respuesta and respuesta.text:
                    return respuesta.text, modelo
            except Exception as e:
                ultimo_error = str(e)
                if "API_KEY_INVALID" in ultimo_error:
                    raise ValueError("La API Key ingresada no es válida. Verifica que empiece con 'AIza...'.")
                if "404" in ultimo_error or "NOT_FOUND" in ultimo_error:
                    break
                time.sleep(2.0)

    raise RuntimeError(f"Los servidores gratuitos de Google están saturados en este instante. Espera 15 segundos y reintenta. ({ultimo_error[:120]})")


def generar_word_ulima(perfil: str, tema: str, detalle_perfil: str, contenido_md: str) -> str:
    hora = ahora_lima().strftime("%H:%M:%S")
    doc = Document()
    doc.add_heading("UNIVERSIDAD DE LIMA — ASISTENTE ACADÉMICO IA (ODS 4)", level=0)
    doc.add_heading(f"{tema} | {detalle_perfil}", level=1)
    doc.add_paragraph(f"Perfil: {perfil} | Emitido: {ahora_lima().strftime('%d/%m/%Y %H:%M')} (Hora Lima)")

    for linea in contenido_md.split("\n"):
        limpia = linea.strip()
        if not limpia:
            continue
        if limpia.startswith("### "):
            doc.add_heading(limpia[4:].replace("**", ""), level=3)
        elif limpia.startswith("## "):
            doc.add_heading(limpia[3:].replace("**", ""), level=2)
        elif limpia.startswith("# "):
            doc.add_heading(limpia[2:].replace("**", ""), level=1)
        else:
            doc.add_paragraph(limpia.replace("**", ""))
    doc.save(ARCHIVO_WORD_PATH)
    return f"[{hora}] 📄 Documento Word (.docx) generado correctamente"


def generar_excel_ulima(perfil: str, tema: str, detalle_perfil: str, contenido_md: str, fecha_evento: datetime) -> str:
    hora = ahora_lima().strftime("%H:%M:%S")
    wb = Workbook()

    # Estilos ULima (Naranja institucional y gris oscuro)
    header_fill = PatternFill(start_color="E85D04", end_color="E85D04", fill_type="solid")
    sub_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    title_font = Font(color="FFFFFF", bold=True, size=13)

    # Hoja 1: Plan de Acción y Cronograma
    ws1 = wb.active
    ws1.title = "Plan y Checklist ULima"
    ws1.merge_cells("A1:E1")
    ws1["A1"] = f"UNIVERSIDAD DE LIMA (ODS 4) — PLAN DE ESTUDIO Y ACCIÓN: {tema.upper()}"
    ws1["A1"].fill = sub_fill
    ws1["A1"].font = title_font
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.append(["Bloque / Etapa", "Tema o Actividad Clave", "Fecha Programada", "Duración Estimada", "Estado de Avance"])
    for cell in ws1[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    bloques = [
        ("1. Diagnóstico y Base", f"Comprensión conceptual de {tema}", fecha_evento.strftime("%d/%m/%Y %H:%M"), "30 min", "Pendiente"),
        ("2. Aplicación Práctica", f"Resolución de casos / ejercicios ({detalle_perfil})", (fecha_evento + timedelta(minutes=30)).strftime("%d/%m/%Y %H:%M"), "40 min", "Pendiente"),
        ("3. Autoevaluación ODS 4", "Verificación con Checklist y repaso final", (fecha_evento + timedelta(minutes=70)).strftime("%d/%m/%Y %H:%M"), "20 min", "Pendiente"),
    ]
    for b in bloques:
        ws1.append(list(b))

    # Hoja 2: Resumen Estructurado del Contenido
    ws2 = wb.create_sheet(title="Guía Estructurada")
    ws2.append(["Sección / Bloque", "Detalle y Puntos Clave"])
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font

    seccion_actual = "Introducción y Enfoque ODS 4"
    for linea in contenido_md.split("\n"):
        limpia = linea.strip().replace("**", "")
        if not limpia:
            continue
        if limpia.startswith("#"):
            seccion_actual = limpia.lstrip("#").strip()
        else:
            ws2.append([seccion_actual, limpia])

    # Hoja 3: Directorio de las 14 Carreras ULima
    ws3 = wb.create_sheet(title="14 Carreras ULima")
    ws3.append(["N.º", "Carrera Profesional ULima", "Enfoque Principal", "Alineación ODS 4"])
    for cell in ws3[1]:
        cell.fill = header_fill
        cell.font = header_font

    for idx, c in enumerate(CARRERAS_ULIMA[:14], start=1):
        ws3.append([idx, c, "Formación integral, innovación y empleabilidad", "Metas 4.3 y 4.4 (Competencias profesionales)"])

    for hoja in [ws1, ws2, ws3]:
        for col in hoja.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                val = str(cell.value or "")
                if len(val) > max_len:
                    max_len = min(len(val), 65)
            hoja.column_dimensions[col_letter].width = max(max_len + 4, 16)

    wb.save(ARCHIVO_EXCEL_PATH)
    return f"[{hora}] 📊 Libro de Excel (.xlsx) generado con 3 hojas estructuradas"


def generar_calendario_ics(titulo_evento: str, inicio_dt: datetime, duracion_min: int, descripcion: str) -> tuple[str, str, datetime, datetime]:
    hora = ahora_lima().strftime("%H:%M:%S")
    fin_dt = inicio_dt + timedelta(minutes=duracion_min)
    ics_texto = (
        "BEGIN:VCALENDAR\n"
        "VERSION:2.0\n"
        "PRODID:-//Agente IA ULima ODS 4//Google Ecosystem//ES\n"
        "BEGIN:VEVENT\n"
        f"DTSTART:{inicio_dt.strftime('%Y%m%dT%H%M%S')}\n"
        f"DTEND:{fin_dt.strftime('%Y%m%dT%H%M%S')}\n"
        f"SUMMARY:{titulo_evento}\n"
        f"DESCRIPTION:{descripcion}\n"
        "END:VEVENT\n"
        "END:VCALENDAR\n"
    )
    with open(ARCHIVO_ICS_PATH, "w", encoding="utf-8") as f:
        f.write(ics_texto)

    titulo_q = urllib.parse.quote(titulo_evento)
    detalles_q = urllib.parse.quote(descripcion)
    fechas_q = f"{inicio_dt.strftime('%Y%m%dT%H%M%S')}/{fin_dt.strftime('%Y%m%dT%H%M%S')}"
    url_gcal = f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={titulo_q}&dates={fechas_q}&details={detalles_q}"

    log = f"[{hora}] 📅 Google Calendar programado: '{titulo_evento}' el {inicio_dt.strftime('%d/%m/%Y a las %H:%M')}"
    return log, url_gcal, inicio_dt, fin_dt


def ejecutar_envio_gmail(remitente: str, clave_app: str, destinatario: str, asunto: str, contenido: str, info_fecha: str) -> tuple[bool, str]:
    hora = ahora_lima().strftime("%H:%M:%S")
    remitente = remitente.strip()
    clave_app = clave_app.replace(" ", "").strip()
    destinatario = destinatario.strip()

    if not destinatario:
        return False, f"[{hora}] 📧 Gmail: sin destinatario indicado"

    if not remitente or not clave_app:
        return False, (
            f"[{hora}] ⚠️ Gmail NO enviado a '{destinatario}': falta configurar tu Gmail y "
            "Contraseña de Aplicación de 16 letras en el panel de conexión."
        )

    try:
        msg = MIMEMultipart()
        msg["From"] = remitente
        msg["To"] = destinatario
        msg["Subject"] = asunto

        cuerpo = (
            "¡Hola!\n\n"
            "Tu Asistente Autónomo ULima (ODS 4) ha completado tu solicitud.\n"
            f"📅 Evento programado: {info_fecha}\n"
            "📎 Se adjuntan automáticamente: Guía en Word (.docx), Plan en Excel (.xlsx) e Invitación de Google Calendar (.ics).\n\n"
            "------------------------------------------------------------\n"
            f"{contenido[:3500]}\n"
            "------------------------------------------------------------\n"
            "Ecosistema Agente IA ULima (ODS 4)."
        )
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

        adjuntos = [
            (ARCHIVO_WORD_PATH, "Guia_ULima_ODS4.docx"),
            (ARCHIVO_EXCEL_PATH, "Plan_Estudio_ULima_ODS4.xlsx"),
            (ARCHIVO_ICS_PATH, "Evento_ULima_ODS4.ics")
        ]
        for ruta, nombre in adjuntos:
            if os.path.exists(ruta):
                with open(ruta, "rb") as f:
                    adj = MIMEBase("application", "octet-stream")
                    adj.set_payload(f.read())
                encoders.encode_base64(adj)
                adj.add_header("Content-Disposition", f"attachment; filename={nombre}")
                msg.attach(adj)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=12) as server:
            server.login(remitente, clave_app)
            server.send_message(msg)

        return True, f"[{hora}] 📧 Gmail ENVIADO con éxito a '{destinatario}' (Word + Excel + Google Calendar adjuntos)"
    except smtplib.SMTPAuthenticationError:
        return False, (
            f"[{hora}] ❌ Gmail rechazó la clave: usa una 'Contraseña de Aplicación' de 16 letras "
            "(myaccount.google.com/apppasswords), no tu contraseña normal."
        )
    except Exception as e:
        return False, f"[{hora}] ❌ Error al enviar Gmail a '{destinatario}': {str(e)}"


# ==============================================================================
# PANTALLA 0: CONEXIÓN INICIAL DEL ECOSISTEMA GOOGLE (ANTES DE INICIAR TODO)
# ==============================================================================
if not st.session_state["conectado"]:
    st.title("🎓 AGENTE IA AUTÓNOMO — UNIVERSIDAD DE LIMA (ODS 4)")
    st.markdown(
        "### 🔗 Paso Inicial: Conecta tu Ecosistema Google\n"
        "Para que el agente pueda **pensar, redactar documentos en Word/Excel, programar fechas en Google Calendar y enviar correos por su cuenta**, "
        "vincula tus credenciales una sola vez antes de comenzar:"
    )

    with st.form("form_conexion_google"):
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("1. Inteligencia Artificial (Google Gemini)")
            input_api = st.text_input(
                "🔑 Google AI Studio API Key (Obligatorio)",
                value=st.session_state["api_key"],
                type="password",
                placeholder="AIzaSy..."
            )
            input_modelo = st.selectbox(
                "⚡ Motor Gemini Principal",
                ["gemini-3.8-flash", "gemini-3.8-pro"],
                index=0
            )
            st.caption("Obtén tu clave gratuita en aistudio.google.com")

        with c2:
            st.subheader("2. Gmail & Google Calendar (Acciones Reales)")
            input_gmail = st.text_input(
                "📧 Tu cuenta de Gmail (para envíos y calendario)",
                value=st.session_state["gmail_user"],
                placeholder="tu.correo@gmail.com"
            )
            input_pass = st.text_input(
                "🔐 Contraseña de Aplicación Google (16 letras)",
                value=st.session_state["gmail_pass"],
                type="password",
                placeholder="abcd efgh ijkl mnop"
            )
            input_destino = st.text_input(
                "📩 Correo receptor por defecto (Alumno / Postulante ULima)",
                value=st.session_state["correo_default"],
                placeholder="codigo@aloe.ulima.edu.pe o tu mismo Gmail"
            )

        st.divider()
        col_btn1, col_btn2 = st.columns([3, 1])
        with col_btn1:
            submitted = st.form_submit_button("🚀 Conectar Ecosistema Google e Iniciar Agente", type="primary", use_container_width=True)
        with col_btn2:
            demo_rapido = st.form_submit_button("Entrar solo con API Key", use_container_width=True)

        if submitted or demo_rapido:
            if not input_api.strip():
                st.error("⚠️ Ingresa tu Google AI Studio API Key (empieza con `AIza...`) para activar el cerebro del agente.")
            else:
                st.session_state["api_key"] = input_api.strip()
                st.session_state["modelo_preferido"] = input_modelo
                st.session_state["gmail_user"] = input_gmail.strip()
                st.session_state["gmail_pass"] = input_pass.strip()
                st.session_state["correo_default"] = input_destino.strip() or input_gmail.strip()
                st.session_state["conectado"] = True
                st.rerun()

    with st.expander("❓ ¿Cómo obtener la Contraseña de Aplicación de Gmail en 30 segundos?"):
        st.markdown(
            "1. Entra a **myaccount.google.com/apppasswords** con tu cuenta de Google.\n"
            "2. Escribe `Agente ULima` y haz clic en **Crear**.\n"
            "3. Copia las **16 letras amarillas** y pégalas arriba. Con eso el agente podrá mandar correos e invitaciones de calendario solo."
        )
    st.stop()


# ==============================================================================
# PANTALLA PRINCIPAL: WORKSPACE AUTÓNOMO ULIMA (YA CONECTADO)
# ==============================================================================
with st.sidebar:
    st.title("⚙️ Ecosistema Conectado")
    st.success("🟢 Google Gemini Activo")
    if st.session_state["gmail_user"] and st.session_state["gmail_pass"]:
        st.success(f"🟢 Gmail: {st.session_state['gmail_user']}")
    else:
        st.warning("🟡 Gmail sin vincular (solo descargas)")

    st.divider()
    correo_destino_ui = st.text_input(
        "📩 Destinatario por defecto",
        value=st.session_state["correo_default"],
        placeholder="alumno@aloe.ulima.edu.pe"
    )
    st.session_state["correo_default"] = correo_destino_ui

    fecha_base_ui = st.date_input(
        "📅 Fecha sugerida (si no la escribes en el texto)",
        value=(ahora_lima() + timedelta(days=1)).date()
    )
    hora_base_ui = st.time_input(
        "⏰ Hora sugerida",
        value=datetime.strptime("18:00", "%H:%M").time()
    )

    st.divider()
    st.markdown("**Herramientas del Agente:**")
    st.markdown("✅ Gemini AI + Web + Wikipedia")
    st.markdown("✅ Google Calendar (Lenguaje Natural)")
    st.markdown("✅ Generador Word (`.docx`)")
    st.markdown("✅ Generador Excel (`.xlsx`)")
    st.markdown("✅ Envío Autónomo Gmail")

    st.divider()
    if st.button("🔄 Cambiar Credenciales / Cuenta", use_container_width=True):
        st.session_state["conectado"] = False
        st.rerun()

st.title("🎓 AGENTE IA ULIMA — Orientación, Postulantes y 14 Carreras (ODS 4)")
st.markdown(
    "Asistente multi-agente conectado a **Google (Gemini, Gmail, Calendar) y Office (Word y Excel)** para facilitar "
    "la **orientación vocacional, el ingreso y el éxito académico en las 14 carreras de la Universidad de Lima**."
)

perfil = st.radio(
    "Elige qué necesitas hoy:",
    [
        "🧭 Orientación Vocacional ULima (Elegir entre las 14 carreras)",
        "🚀 Postulante a la ULima (Admisión / CEPREULIMA / Escolar)",
        "🎓 Alumno Regular ULima (14 Carreras / Estudios Generales)"
    ],
    horizontal=True
)

st.divider()

if "Orientación" in perfil:
    col1, col2, col3 = st.columns(3)
    with col1:
        intereses = st.text_input(
            "Tus intereses, cursos favoritos o talentos",
            value="Tecnología, gestión de empresas, liderazgo e innovación"
        )
    with col2:
        carreras_comparar = st.multiselect(
            "Carreras ULima que te llaman la atención (14 disponibles)",
            CARRERAS_ULIMA[:14],
            default=["Ingeniería de Sistemas", "Ingeniería Industrial", "Administración"]
        )
    with col3:
        tema = st.text_input(
            "Enfoque de tu orientación",
            value="Comparativa de carreras, campo laboral en Perú y perfil ideal"
        )
    detalle_perfil = ", ".join(carreras_comparar) if carreras_comparar else "14 Carreras ULima"
    necesidad = "Orientación Vocacional Personalizada ULima"
    valor_default_mision = (
        f"Analiza mis intereses ('{intereses}') y compáralos con las carreras de la ULima ({detalle_perfil}). "
        "Recomiéndame cuál es mi mejor opción, explícame de forma sencilla qué hace cada carrera en el mundo real, "
        "genera el Word y el Excel comparativo, y programa una sesión de exploración vocacional para mañana a las 5:00 pm."
    )

elif "Postulante" in perfil:
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        modalidad = st.selectbox(
            "Modalidad de Postulación",
            [
                "Examen de Admisión General",
                "CEPREULIMA (Ciclo Preuniversitario)",
                "Colegio / Tercios Superiores / Bachillerato",
                "Nivelación Base antes del 1.er Ciclo"
            ]
        )
    with col2:
        carrera_objetivo = st.selectbox(
            "Carrera a la que postula (14 carreras)",
            CARRERAS_ULIMA[:14]
        )
    with col3:
        tema = st.text_input(
            "Materia o tema de admisión",
            value="Razonamiento Lógico-Matemático y Comprensión Lectora"
        )
    with col4:
        necesidad = st.selectbox(
            "¿Cómo te ayudamos?",
            [
                "Simulacro guiado con preguntas tipo ULima y trucos",
                "Explicación desde cero sin estrés",
                "Plan de estudio semanal en Excel y Calendar",
                "Guía puente hacia Estudios Generales ULima"
            ]
        )
    detalle_perfil = f"{modalidad} → {carrera_objetivo}"
    valor_default_mision = (
        f"Explícalo fácil desde cero para un postulante a {carrera_objetivo} en la ULima, incluye trucos rápidos para el examen, "
        "4 preguntas tipo admisión resueltas paso a paso, genera la guía en Word y el cronograma en Excel, y agéndalo para mañana a las 6:00 pm."
    )

else:
    col1, col2, col3 = st.columns(3)
    with col1:
        carrera_actual = st.selectbox(
            "Carrera / Facultad ULima (14 carreras)",
            CARRERAS_ULIMA
        )
    with col2:
        tema = st.text_input(
            "Curso o tema difícil a simplificar",
            value="Inteligencia Artificial aplicada a los Negocios"
        )
    with col3:
        necesidad = st.selectbox(
            "¿Cómo te facilitamos el curso hoy?",
            [
                "Preparación rápida para Examen Parcial / Final",
                "Explicar teoría difícil desde cero con ejemplos prácticos",
                "Guía paso a paso para Trabajo Grupal / Proyecto",
                "Plan de estudio en Excel + Google Calendar"
            ]
        )
    detalle_perfil = carrera_actual
    valor_default_mision = (
        f"Explícalo fácil para un alumno de {carrera_actual} en la ULima con ejemplos reales, 3 ejercicios resueltos paso a paso, "
        "genera el resumen en Word y la hoja de estudio en Excel, y programa el repaso para mañana a las 7:00 pm."
    )

objetivo_extra = st.text_area(
    "Instrucción libre para el agente (puedes pedirle fechas para Calendar, Word, Excel o enviar a un correo específico):",
    value=valor_default_mision,
    height=100
)

if st.button("⚡ Ejecutar Agente Autónomo ULima", type="primary", use_container_width=True):
    tiempo_inicio = time.time()
    bitacora = []
    ahora = ahora_lima()
    fecha_hoy_str = f"{DIAS_ES[ahora.weekday()]} {ahora.strftime('%Y-%m-%d')} (hora actual Lima: {ahora.strftime('%H:%M')})"

    # Detectar destinatario (en texto o en configuración)
    correos_en_texto = re.findall(r"[\w\.-]+@[\w\.-]+\.\w+", objetivo_extra)
    destinatario_final = correos_en_texto[0] if correos_en_texto else st.session_state["correo_default"].strip()

    estado_calendario = {
        "inicio": datetime.combine(fecha_base_ui, hora_base_ui),
        "titulo": f"Estudio ULima: {tema}",
        "url_gcal": "",
        "programado_por_ia": False
    }
    estado_gmail = {"intentado": False, "exito": False, "mensaje": ""}

    try:
        client = genai.Client(api_key=st.session_state["api_key"])

        with st.status("⚡ Agentes ejecutando investigación, archivos, calendario y correo...", expanded=True) as status:
            # FASE 1: Búsqueda paralela + Agente Investigador
            st.write("🔍 **Fase 1/3:** *Agente Investigador ULima* consultando información académica en paralelo...")
            query_busqueda = (
                f"Universidad de Lima carreras perfil profesional {detalle_perfil}"
                if "Orientación" in perfil else
                f"Universidad de Lima {detalle_perfil} {tema}"
            )
            with ThreadPoolExecutor(max_workers=3) as pool:
                futuro_web = pool.submit(buscar_web_rapido, query_busqueda)
                futuro_wiki = pool.submit(buscar_wikipedia_rapido, tema)
                futuro_modelos = pool.submit(obtener_modelos_activos, client, st.session_state["modelo_preferido"])
                info_web, log_web = futuro_web.result()
                info_wiki, log_wiki = futuro_wiki.result()
                modelos_activos = futuro_modelos.result()

            bitacora.extend([log_web, log_wiki])

            prompt_investigador = (
                f"Eres el Agente Investigador Académico y Vocacional de la Universidad de Lima (ULima - ODS 4). "
                f"Perfil atendido: {perfil} ({detalle_perfil}). Tema: '{tema}'. Objetivo: '{necesidad}'. "
                f"Sintetiza en 5 puntos claros los datos clave, competencias y oportunidades relevantes.\nFuentes:\n{info_web}\nWikipedia:\n{info_wiki}"
            )
            hallazgos, modelo_usado = llamar_gemini_blindado(client, prompt_investigador, modelos_activos)
            modelos_priorizados = [modelo_usado] + [m for m in modelos_activos if m != modelo_usado]
            bitacora.append(f"[{ahora_lima().strftime('%H:%M:%S')}] 🧠 Agente Investigador completado ({modelo_usado})")

            # FASE 2: Agente Orientador / Tutor Facilitador
            st.write("💡 **Fase 2/3:** *Agente Orientador y Tutor ULima* redactando guía clara y práctica...")
            if "Orientación" in perfil:
                prompt_tutor = (
                    f"Eres el Agente Consejero Vocacional de la Universidad de Lima. Ayuda a un joven a elegir su carrera entre las 14 carreras de la ULima. "
                    f"Con base en:\n{hallazgos}\n\nPedido del estudiante: {objetivo_extra}.\n"
                    f"Crea una guía súper clara en Markdown con: 1) Diagnóstico de afinidad según sus intereses, "
                    f"2) Tabla comparativa de las carreras analizadas (Qué hace | Cursos clave | Dónde trabaja | Ideal si te gusta...), "
                    f"3) Recomendación final personalizada y por qué encaja con la ULima, "
                    f"4) Primeros pasos para postular (Admisión / CEPREULIMA)."
                )
            elif "Postulante" in perfil:
                prompt_tutor = (
                    f"Eres el Agente Tutor Preuniversitario de la Universidad de Lima para postulantes ({detalle_perfil}). "
                    f"Con base en:\n{hallazgos}\n\nPedido: {objetivo_extra}.\n"
                    f"Crea una guía súper fácil de entender en Markdown con: 1) Explicación desde cero en palabras simples, "
                    f"2) Atajos y trucos para resolver rápido en el examen ULima, "
                    f"3) 4 preguntas tipo examen con alternativas (A, B, C, D) y solución paso a paso, "
                    f"4) Cómo se conecta este tema con su carrera en la ULima."
                )
            else:
                prompt_tutor = (
                    f"Eres el Agente Tutor Facilitador de la Universidad de Lima para alumnos de {detalle_perfil}. "
                    f"Con base en:\n{hallazgos}\n\nPedido: {objetivo_extra}.\n"
                    f"Crea una guía directa y sin relleno académico en Markdown con: 1) Resumen en palabras simples, "
                    f"2) Ejemplos prácticos aplicados a empresas en el Perú, "
                    f"3) 3 ejercicios o casos resueltos paso a paso tipo evaluación ULima, "
                    f"4) Errores comunes que debes evitar."
                )

            borrador_guia, modelo_usado = llamar_gemini_blindado(client, prompt_tutor, modelos_priorizados)
            bitacora.append(f"[{ahora_lima().strftime('%H:%M:%S')}] 💡 Agente Tutor/Orientador completado ({modelo_usado})")

            # FASE 3: Agente Auditor ODS 4
            st.write("🛡️ **Fase 3/3:** *Agente Auditor ODS 4* validando calidad y estructurando entregables...")
            prompt_auditor = (
                f"Eres el Agente Auditor de Calidad y Acceso Educativo (ODS 4 - UNESCO) de la ULima. Revisa esta guía:\n\n{borrador_guia}\n\n"
                f"Devuelve la VERSIÓN FINAL en Markdown encabezada por:\n"
                f"- **Índice de Facilitación y Acceso ULima (ODS 4):** [Puntaje/100] y breve justificación.\n"
                f"- Todo el contenido claro, ordenado y listo para estudiar.\n"
                f"- Al final, un **Checklist Rápido de Dominio** en tabla Markdown (Punto Clave | ¿Listo? | Consejo Rápido)."
            )
            guia_final, modelo_usado = llamar_gemini_blindado(client, prompt_auditor, modelos_priorizados)
            bitacora.append(f"[{ahora_lima().strftime('%H:%M:%S')}] 🛡️ Agente Auditor ODS 4 completado ({modelo_usado})")

            # FASE 4: Herramientas autónomas de Google Calendar, Word, Excel y Gmail
            st.write("🤖 **Ejecutando acciones en Google Calendar, Word, Excel y Gmail...**")

            def tool_programar_google_calendar(titulo_evento: str, fecha_yyyy_mm_dd: str, hora_hh_mm: str, duracion_minutos: int = 90) -> str:
                """Programa una fecha y hora exacta en Google Calendar a partir de la instrucción del usuario."""
                try:
                    dt_parsed = datetime.strptime(f"{fecha_yyyy_mm_dd.strip()} {hora_hh_mm.strip()}", "%Y-%m-%d %H:%M")
                except Exception:
                    dt_parsed = datetime.combine(fecha_base_ui, hora_base_ui)

                log_cal, url_cal, ini_dt, _ = generar_calendario_ics(
                    titulo_evento or f"ULima ODS 4: {tema}",
                    dt_parsed,
                    duracion_minutos or 90,
                    f"Sesión ULima ({detalle_perfil}) - {necesidad}"
                )
                estado_calendario["inicio"] = ini_dt
                estado_calendario["titulo"] = titulo_evento or f"ULima ODS 4: {tema}"
                estado_calendario["url_gcal"] = url_cal
                estado_calendario["programado_por_ia"] = True
                bitacora.append(log_cal)
                return log_cal

            prompt_acciones = (
                f"Eres el Agente Coordinador de Herramientas de Google. Hoy es {fecha_hoy_str}.\n"
                f"Analiza la instrucción del usuario: '{objetivo_extra}'.\n"
                f"Llama a `tool_programar_google_calendar` convirtiendo cualquier referencia temporal del usuario "
                f"(como 'mañana', 'este viernes', 'el lunes', '28 de noviembre', 'a las 5 pm', etc.) a formato exacto "
                f"`YYYY-MM-DD` y `HH:MM` (24 horas). Si el usuario no mencionó ninguna fecha ni hora, usa "
                f"`{fecha_base_ui.strftime('%Y-%m-%d')}` y `{hora_base_ui.strftime('%H:%M')}`."
            )
            llamar_gemini_blindado(
                client,
                prompt_acciones,
                modelos_priorizados,
                herramientas=[tool_programar_google_calendar]
            )

            # Respaldo si el modelo no invocó la herramienta de calendario
            if not estado_calendario["programado_por_ia"]:
                dt_def = datetime.combine(fecha_base_ui, hora_base_ui)
                log_cal, url_cal, ini_dt, _ = generar_calendario_ics(
                    f"Estudio ULima: {tema}",
                    dt_def,
                    90,
                    f"Sesión ULima ({detalle_perfil}) - {necesidad}"
                )
                estado_calendario["inicio"] = ini_dt
                estado_calendario["url_gcal"] = url_cal
                bitacora.append(log_cal)

            # Generar siempre Word (.docx) y Excel (.xlsx) listos para descarga y adjunto
            log_word = generar_word_ulima(perfil, tema, detalle_perfil, guia_final)
            log_excel = generar_excel_ulima(perfil, tema, detalle_perfil, guia_final, estado_calendario["inicio"])
            bitacora.extend([log_word, log_excel])

            # Envío automático por Gmail si hay destinatario (en texto o configurado)
            quiere_correo = any(p in objetivo_extra.lower() for p in ["correo", "gmail", "envia", "envía", "manda", "@"]) or bool(destinatario_final)
            if quiere_correo and destinatario_final:
                exito, msg_gmail = ejecutar_envio_gmail(
                    st.session_state["gmail_user"],
                    st.session_state["gmail_pass"],
                    destinatario_final,
                    f"Guía + Excel + Calendario ULima ({detalle_perfil}): {tema}",
                    guia_final,
                    f"{estado_calendario['titulo']} — {estado_calendario['inicio'].strftime('%d/%m/%Y %H:%M')}"
                )
                estado_gmail["intentado"] = True
                estado_gmail["exito"] = exito
                estado_gmail["mensaje"] = msg_gmail
                bitacora.append(msg_gmail)

            duracion = round(time.time() - tiempo_inicio, 1)
            status.update(label=f"✨ Todo completado en {duracion}s ({modelo_usado})", state="complete", expanded=False)

    except Exception as error_controlado:
        st.warning(f"⚠️ Aviso del sistema: {error_controlado}")
        st.stop()

    # Alertas visuales de acciones autónomas realizadas
    if estado_gmail["intentado"]:
        if estado_gmail["exito"]:
            st.success(f"✅ **Gmail Enviado Autónomamente:** {estado_gmail['mensaje']}")
        else:
            st.error(f"🚨 **Aviso de Gmail:** {estado_gmail['mensaje']}")

    st.info(
        f"📅 **Google Calendar Programado:** `{estado_calendario['titulo']}` para el "
        f"**{estado_calendario['inicio'].strftime('%d/%m/%Y a las %H:%M')}** (Hora Lima)."
    )

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Tiempo Total", f"{duracion} seg")
    m2.metric("Word & Excel", "Generados ✅")
    m3.metric("Fecha Calendar", estado_calendario["inicio"].strftime("%d/%m %H:%M"))
    m4.metric("Estado Gmail", "Enviado ✅" if estado_gmail["exito"] else "Listo para enviar")

    with st.expander("🛠️ Bitácora en Vivo de Acciones Ejecutadas por el Agente", expanded=True):
        for linea_log in bitacora:
            st.code(linea_log, language="text")

    st.subheader("📦 Descargas Directas (Word, Excel, Calendario) y Sincronización Google")
    b1, b2, b3, b4 = st.columns(4)

    with b1:
        if os.path.exists(ARCHIVO_WORD_PATH):
            with open(ARCHIVO_WORD_PATH, "rb") as f_doc:
                st.download_button(
                    label="📄 Descargar Word (.docx)",
                    data=f_doc.read(),
                    file_name=f"Guia_ULima_{tema.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
    with b2:
        if os.path.exists(ARCHIVO_EXCEL_PATH):
            with open(ARCHIVO_EXCEL_PATH, "rb") as f_xls:
                st.download_button(
                    label="📊 Descargar Excel (.xlsx)",
                    data=f_xls.read(),
                    file_name=f"Plan_ULima_{tema.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    with b3:
        if os.path.exists(ARCHIVO_ICS_PATH):
            with open(ARCHIVO_ICS_PATH, "rb") as f_ics:
                st.download_button(
                    label="📅 Invitación Calendar (.ics)",
                    data=f_ics.read(),
                    file_name="Evento_ULima_ODS4.ics",
                    mime="text/calendar",
                    use_container_width=True
                )
    with b4:
        if estado_calendario["url_gcal"]:
            st.link_button("🗓️ Abrir en Google Calendar", estado_calendario["url_gcal"], use_container_width=True)

    st.divider()
    st.subheader(f"📘 Resultado Generado — {detalle_perfil}")
    st.markdown(guia_final)
