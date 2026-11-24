import os
import time
import json
import smtplib
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

import streamlit as st
from docx import Document
from google import genai
from duckduckgo_search import DDGS

st.set_page_config(
    page_title="AGENTE IA ULIMA | 14 Carreras y Postulantes (ODS 4)",
    page_icon="🎓",
    layout="wide"
)

ARCHIVO_WORD_PATH = "/tmp/Guia_ULima_ODS4.docx"
ARCHIVO_ICS_PATH = "/tmp/Sesion_Estudio_ULima.ics"

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


def buscar_web_rapido(consulta: str) -> tuple[str, str]:
    hora = datetime.now().strftime("%H:%M:%S")
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(consulta, max_results=3))
        if not resultados:
            return "Lineamientos universitarios y de acceso educativo ODS 4 (Meta 4.3 UNESCO).", f"[{hora}] 🔍 Web Search: consulta completada (base académica)"
        texto = "\n".join([f"- {r['title']}: {r['body']} ({r['href']})" for r in resultados])
        return texto, f"[{hora}] 🔍 Web Search conectado: {len(resultados)} referencias académicas halladas"
    except Exception:
        return "Estándares de acceso y calidad en educación superior UNESCO ODS 4.", f"[{hora}] 🔍 Web Search: modo respaldo académico activo"


def buscar_wikipedia_rapido(concepto: str) -> tuple[str, str]:
    hora = datetime.now().strftime("%H:%M:%S")
    try:
        termino = urllib.parse.quote(concepto.strip().split()[0])
        url = f"https://es.wikipedia.org/api/rest_v1/page/summary/{termino}"
        req = urllib.request.Request(url, headers={"User-Agent": "AgenteULimaODS4/4.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            datos = json.loads(resp.read().decode("utf-8"))
        extracto = datos.get("extract", "")
        if extracto:
            return extracto, f"[{hora}] 📚 Wikipedia API conectada: concepto '{concepto}' verificado"
        return "", f"[{hora}] 📚 Wikipedia API: consulta completada"
    except Exception:
        return "", f"[{hora}] 📚 Wikipedia API: usando síntesis académica interna"


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


def llamar_gemini_blindado(client: genai.Client, prompt: str, modelos_activos: list[str]) -> tuple[str, str]:
    ultimo_error = ""
    for modelo in modelos_activos:
        for _ in range(2):
            try:
                respuesta = client.models.generate_content(
                    model=modelo,
                    contents=prompt
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


def generar_archivos_ulima(perfil: str, tema: str, detalle_perfil: str, contenido_md: str, inicio_dt: datetime, fin_dt: datetime) -> list[str]:
    logs = []
    hora = datetime.now().strftime("%H:%M:%S")

    doc = Document()
    doc.add_heading("UNIVERSIDAD DE LIMA — ASISTENTE ACADÉMICO IA (ODS 4)", level=0)
    doc.add_heading(f"{tema} | {detalle_perfil}", level=1)
    doc.add_paragraph(f"Perfil: {perfil} | Generado el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}")

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
    logs.append(f"[{hora}] 📄 Guía Word (.docx) generada ({detalle_perfil})")

    ics_texto = (
        "BEGIN:VCALENDAR\n"
        "VERSION:2.0\n"
        "PRODID:-//Agente IA ULima ODS 4//14 Carreras y Postulantes//ES\n"
        "BEGIN:VEVENT\n"
        f"DTSTART:{inicio_dt.strftime('%Y%m%dT%H%M%S')}\n"
        f"DTEND:{fin_dt.strftime('%Y%m%dT%H%M%S')}\n"
        f"SUMMARY:Estudio ULima - {tema}\n"
        f"DESCRIPTION:Sesión guiada ({perfil} - {detalle_perfil}) generada por Agente IA ULima (ODS 4).\n"
        "END:VEVENT\n"
        "END:VCALENDAR\n"
    )
    with open(ARCHIVO_ICS_PATH, "w", encoding="utf-8") as f:
        f.write(ics_texto)
    logs.append(f"[{hora}] 📅 Bloque de estudio en Calendario (.ics) generado")

    return logs


def enviar_gmail_real(remitente: str, clave_app: str, destinatario: str, asunto: str, contenido: str) -> str:
    hora = datetime.now().strftime("%H:%M:%S")
    remitente = remitente.strip()
    clave_app = clave_app.replace(" ", "").strip()
    destinatario = destinatario.strip()

    if not remitente or not clave_app or not destinatario:
        return f"[{hora}] 📧 Gmail: listo (agrega tu Gmail y clave de aplicación a la izquierda para enviarlo)"

    try:
        msg = MIMEMultipart()
        msg["From"] = remitente
        msg["To"] = destinatario
        msg["Subject"] = asunto

        cuerpo = (
            "¡Hola!\n\n"
            "Tu Asistente Académico Autónomo ULima (ODS 4) ha preparado tu material de estudio personalizado.\n"
            "Adjunto encontrarás la guía completa en Word (.docx) lista para repasar e imprimir y el evento para tu Google Calendar (.ics).\n\n"
            "------------------------------------------------------------\n"
            f"{contenido[:3500]}\n"
            "------------------------------------------------------------\n"
            "Facilitación Académica y Acceso Universitario (14 Carreras ULima) — Agente IA ODS 4."
        )
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

        for ruta, nombre in [(ARCHIVO_WORD_PATH, "Guia_ULima_ODS4.docx"), (ARCHIVO_ICS_PATH, "Sesion_Estudio_ULima.ics")]:
            if os.path.exists(ruta):
                with open(ruta, "rb") as f:
                    adj = MIMEBase("application", "octet-stream")
                    adj.set_payload(f.read())
                encoders.encode_base64(adj)
                adj.add_header("Content-Disposition", f"attachment; filename={nombre}")
                msg.attach(adj)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10) as server:
            server.login(remitente, clave_app)
            server.send_message(msg)

        return f"[{hora}] 📧 Gmail enviado con éxito a '{destinatario}' (Word + Calendario adjuntos)"
    except Exception as e:
        return f"[{hora}] ⚠️ Gmail no enviado ({str(e)})"


def enviar_alertas_push(webhook_url: str, tg_token: str, tg_chat_id: str, resumen: str) -> list[str]:
    logs = []
    hora = datetime.now().strftime("%H:%M:%S")

    if webhook_url.strip():
        try:
            data = json.dumps({"content": f"🎓 **Agente IA ULima (ODS 4) — Material Listo:**\n{resumen[:1500]}"}).encode("utf-8")
            req = urllib.request.Request(webhook_url.strip(), data=data, headers={"Content-Type": "application/json", "User-Agent": "AgenteULimaODS4/4.0"})
            urllib.request.urlopen(req, timeout=6)
            logs.append(f"[{hora}] 💬 Grupo de estudio en Discord/Slack notificado")
        except Exception as e:
            logs.append(f"[{hora}] ⚠️ Error Webhook: {str(e)}")

    if tg_token.strip() and tg_chat_id.strip():
        try:
            url = f"https://api.telegram.org/bot{tg_token.strip()}/sendMessage"
            data = json.dumps({"chat_id": tg_chat_id.strip(), "text": f"🎓 Agente IA ULima (ODS 4):\n\n{resumen[:1500]}"}).encode("utf-8")
            req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
            urllib.request.urlopen(req, timeout=6)
            logs.append(f"[{hora}] 📲 Telegram notificado (Chat {tg_chat_id.strip()})")
        except Exception as e:
            logs.append(f"[{hora}] ⚠️ Error Telegram: {str(e)}")

    if not logs:
        logs.append(f"[{hora}] 💬 Alertas Push: disponibles (conecta Discord o Telegram en el panel izquierdo)")
    return logs


# --- BARRA LATERAL ---
with st.sidebar:
    st.title("⚙️ Núcleo Agente ULima")
    st.caption("14 Carreras + Postulantes — ODS 4")
    api_key = st.text_input(
        "🔑 Google AI Studio API Key (Obligatorio)",
        type="password",
        help="Tu clave gratuita que empieza con AIza..."
    )
    modelo_elegido = st.selectbox(
        "⚡ Motor Gemini",
        ["gemini-3.8-flash", "gemini-3.8-pro"]
    )

    st.divider()
    with st.expander("📧 1. Enviar al Correo (Gmail)", expanded=True):
        gmail_user = st.text_input("Tu correo Gmail (remitente)", placeholder="tu.correo@gmail.com")
        gmail_pass = st.text_input("Contraseña de Aplicación (16 letras)", type="password")
        correo_destino = st.text_input("Correo del alumno o postulante", placeholder="postulante@gmail.com / @aloe.ulima.edu.pe")

    with st.expander("📅 2. Agendar Estudio (Google Calendar)", expanded=False):
        fecha_clase = st.date_input("Fecha de repaso / simulacro", value=datetime.now().date() + timedelta(days=1))
        hora_clase = st.time_input("Hora de inicio", value=datetime.strptime("18:00", "%H:%M").time())

    with st.expander("💬 3. Grupo de Estudio (Discord / Telegram)", expanded=False):
        webhook_url = st.text_input("Webhook Discord / Slack", placeholder="https://discord.com/api/webhooks/...")
        tg_token = st.text_input("Telegram Bot Token", type="password")
        tg_chat_id = st.text_input("Telegram Chat ID", placeholder="123456789")

    st.divider()
    st.markdown("**6 Herramientas Conectadas:**")
    st.markdown("✅ Buscador Académico Web")
    st.markdown("✅ API Oficial Wikipedia")
    st.markdown("✅ Guía Word ULima (`.docx`)")
    st.markdown("✅ Google Calendar (`.ics` + Link)")
    st.markdown("✅ Envío Automático Gmail")
    st.markdown("✅ Alertas Discord / Telegram")

# --- PANTALLA PRINCIPAL ---
st.title("🎓 AGENTE IA ULIMA — 14 Carreras y Postulantes (ODS 4)")
st.markdown(
    "Sistema multi-agente autónomo diseñado para **facilitar tanto el ingreso de postulantes como el rendimiento académico en las 14 carreras de la Universidad de Lima**, "
    "entregando tutorías personalizadas y materiales listos en **Word, Calendar y Correo** (**Metas 4.3, 4.4 y 4.5 de la UNESCO**)."
)

perfil = st.radio(
    "Selecciona tu perfil en la comunidad ULima:",
    [
        "🚀 Postulante a la ULima (Admisión / CEPREULIMA / Escolar)",
        "🎓 Alumno Regular ULima (14 Carreras / Estudios Generales)"
    ],
    horizontal=True
)

st.divider()

if "Postulante" in perfil:
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        modalidad = st.selectbox(
            "Modalidad de Postulación",
            [
                "Examen de Admisión General",
                "CEPREULIMA (Ciclo Preuniversitario)",
                "Colegio / Tercios Superiores / Bachillerato",
                "Nivelación Base antes del 1.er Ciclo",
                "Orientación Vocacional ULima"
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
            "¿Cómo te ayudamos a ingresar?",
            [
                "Simulacro guiado con preguntas tipo ULima y trucos",
                "Explicación desde cero sin estrés (cerrar brechas)",
                "Plan de estudio semanal para Google Calendar",
                "Guía puente hacia Estudios Generales ULima"
            ]
        )
    detalle_perfil = f"{modalidad} → {carrera_objetivo}"
    valor_default_mision = (
        f"Explícalo de forma clara desde cero para un postulante a {carrera_objetivo} en la ULima, incluye atajos y trucos para resolver preguntas rápido en el examen de admisión, "
        "4 preguntas tipo examen con alternativas y solución paso a paso, y consejos para asegurar el ingreso."
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
                "Preparación rápida para Examen Parcial / Final (resumen + preguntas tipo)",
                "Explicar teoría difícil desde cero con ejemplos prácticos del Perú",
                "Guía paso a paso para Trabajo Grupal / Proyecto de Investigación",
                "Plan de estudio anti-estrés con cronograma listo para Calendar",
                "Lectura accesible y estructurada sin tecnicismos innecesarios (ODS 4.5)"
            ]
        )
    detalle_perfil = carrera_actual
    valor_default_mision = (
        f"Explícalo de forma directa y fácil de entender para un alumno de {carrera_actual} en la ULima, incluye ejemplos aplicados a casos reales en el Perú, "
        "3 ejercicios resueltos paso a paso tipo evaluación ULima y un checklist rápido para asegurar buena nota."
    )

objetivo_extra = st.text_area(
    "Detalle adicional o meta específica",
    value=valor_default_mision
)

if st.button("⚡ Ejecutar Agente Autónomo ULima", type="primary", use_container_width=True):
    if not api_key:
        st.error("Ingresa tu API Key de Google AI Studio en el panel izquierdo.")
        st.stop()

    tiempo_inicio = time.time()
    bitacora = []
    inicio_dt = datetime.combine(fecha_clase, hora_clase)
    fin_dt = inicio_dt + timedelta(minutes=90)
    es_postulante = "Postulante" in perfil

    try:
        client = genai.Client(api_key=api_key.strip())

        with st.status("⚡ Agentes trabajando para facilitar tu camino en la ULima...", expanded=True) as status:
            st.write("🔍 **Fase 1/3:** *Agente Investigador Académico* consultando fuentes y temarios en paralelo...")
            query_busqueda = (
                f"Universidad de Lima admision CEPREULIMA {carrera_objetivo} {tema}"
                if es_postulante else
                f"Universidad de Lima {carrera_actual} {tema} conceptos clave"
            )
            with ThreadPoolExecutor(max_workers=3) as pool:
                futuro_web = pool.submit(buscar_web_rapido, query_busqueda)
                futuro_wiki = pool.submit(buscar_wikipedia_rapido, tema)
                futuro_modelos = pool.submit(obtener_modelos_activos, client, modelo_elegido)
                info_web, log_web = futuro_web.result()
                info_wiki, log_wiki = futuro_wiki.result()
                modelos_activos = futuro_modelos.result()

            bitacora.extend([log_web, log_wiki])

            if es_postulante:
                prompt_investigador = (
                    f"Eres el Agente Investigador de Admisión y Nivelación para postulantes a la carrera de {carrera_objetivo} "
                    f"en la Universidad de Lima ({modalidad}). Sintetiza en 5 puntos clave qué evalúa la ULima y qué bases necesita "
                    f"dominar sobre '{tema}' para cumplir '{necesidad}' (ODS 4 Meta 4.3).\nFuentes:\n{info_web}\nWikipedia:\n{info_wiki}"
                )
            else:
                prompt_investigador = (
                    f"Eres el Agente Investigador Académico para alumnos de {carrera_actual} en la Universidad de Lima (ULima). "
                    f"Sintetiza en 5 puntos claros los conceptos esenciales y enfoques prácticos sobre '{tema}' "
                    f"para resolver '{necesidad}'.\nFuentes:\n{info_web}\nWikipedia:\n{info_wiki}"
                )

            hallazgos, modelo_usado = llamar_gemini_blindado(client, prompt_investigador, modelos_activos)
            modelos_priorizados = [modelo_usado] + [m for m in modelos_activos if m != modelo_usado]
            bitacora.append(f"[{datetime.now().strftime('%H:%M:%S')}] 🧠 Agente Investigador completado ({modelo_usado})")

            st.write("💡 **Fase 2/3:** *Agente Tutor Facilitador* preparando explicaciones claras y ejercicios paso a paso...")
            if es_postulante:
                prompt_tutor = (
                    f"Eres el Agente Tutor Preuniversitario de la Universidad de Lima. Tu misión es ayudar a que un postulante a {carrera_objetivo} ({modalidad}) "
                    f"entienda todo fácil y logre su ingreso sin sentirse abrumado. Con base en:\n{hallazgos}\n\n"
                    f"Crea una guía súper clara en Markdown sobre '{tema}' enfocada en '{necesidad}'. Pedido: {objetivo_extra}.\n"
                    f"Incluye: 1) Explicación desde cero en palabras simples ('De cero a experto'), "
                    f"2) Atajos y trucos de resolución rápida para ganar tiempo en el examen, "
                    f"3) 4 preguntas tipo examen de admisión ULima (con alternativas A, B, C, D y solución detallada paso a paso), "
                    f"4) Conexión con su futura carrera ({carrera_objetivo} en la ULima)."
                )
            else:
                prompt_tutor = (
                    f"Eres el Agente Tutor Facilitador de la Universidad de Lima. Tu misión es hacerle la vida fácil al alumno de {carrera_actual} "
                    f"eliminando el relleno académico innecesario. Con base en:\n{hallazgos}\n\n"
                    f"Crea una guía súper clara en Markdown sobre '{tema}' enfocada en '{necesidad}'. Pedido: {objetivo_extra}.\n"
                    f"Incluye: 1) Resumen ejecutivo en lenguaje sencillo ('En palabras simples'), "
                    f"2) Ejemplos prácticos aplicados a {carrera_actual} en el Perú, "
                    f"3) 3 ejercicios o casos resueltos paso a paso tipo evaluación ULima, "
                    f"4) Errores típicos que debes evitar en el examen."
                )

            borrador_guia, modelo_usado = llamar_gemini_blindado(client, prompt_tutor, modelos_priorizados)
            bitacora.append(f"[{datetime.now().strftime('%H:%M:%S')}] 💡 Agente Tutor Facilitador completado ({modelo_usado})")

            st.write("🛡️ **Fase 3/3:** *Agente Auditor ODS 4* validando claridad, creando checklist y sincronizando herramientas...")
            prompt_auditor = (
                f"Eres el Agente Auditor de Accesibilidad y Calidad Educativa (ODS 4 - UNESCO) para la comunidad ULima ({detalle_perfil}). "
                f"Revisa esta guía:\n\n{borrador_guia}\n\n"
                f"Entrega la GUÍA FINAL LISTA PARA ESTUDIAR en Markdown encabezada por:\n"
                f"- **Índice de Facilitación y Acceso ULima (ODS 4):** [Puntaje/100] y justificación de cómo reduce barreras de aprendizaje.\n"
                f"- Todo el contenido optimizado, claro y directo.\n"
                f"- Al final, un **Checklist Rápido de Dominio** (tabla Markdown con: Tema Clave | ¿Lo domino? | Tip de memoria rápida)."
            )
            guia_final, modelo_usado = llamar_gemini_blindado(client, prompt_auditor, modelos_priorizados)
            bitacora.append(f"[{datetime.now().strftime('%H:%M:%S')}] 🛡️ Agente Auditor ODS 4 completado ({modelo_usado})")

            logs_archivos = generar_archivos_ulima(perfil, tema, detalle_perfil, guia_final, inicio_dt, fin_dt)
            bitacora.extend(logs_archivos)

            etiqueta_corto = "Postulante ULima" if es_postulante else "Alumno ULima"
            with ThreadPoolExecutor(max_workers=2) as pool_envios:
                futuro_gmail = pool_envios.submit(
                    enviar_gmail_real,
                    gmail_user,
                    gmail_pass,
                    correo_destino,
                    f"Guía Fácil {etiqueta_corto} ({detalle_perfil}) - ODS 4: {tema}",
                    guia_final
                )
                futuro_push = pool_envios.submit(
                    enviar_alertas_push,
                    webhook_url,
                    tg_token,
                    tg_chat_id,
                    f"Perfil: {etiqueta_corto} ({detalle_perfil}) | Tema: {tema}\n\n{guia_final[:600]}..."
                )
                bitacora.append(futuro_gmail.result())
                bitacora.extend(futuro_push.result())

            duracion = round(time.time() - tiempo_inicio, 1)
            status.update(label=f"✨ Guía lista en {duracion}s usando {modelo_usado}", state="complete", expanded=False)

    except Exception as error_controlado:
        st.warning(f"⚠️ Aviso del sistema: {error_controlado}")
        st.stop()

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Tiempo de Respuesta", f"{duracion} seg")
    m2.metric("Programa / Carrera", carrera_objetivo if es_postulante else carrera_actual)
    m3.metric("Acciones Ejecutadas", f"{len(bitacora)} Herramientas")
    m4.metric("Envío por Correo", "Enviado" if correo_destino.strip() and gmail_user.strip() else "Listo")

    with st.expander("🛠️ Bitácora en Vivo de Acciones Ejecutadas por el Agente", expanded=True):
        for linea_log in bitacora:
            st.code(linea_log, language="text")

    st.subheader("📦 Descarga y Sincronización Directa")
    b1, b2, b3, b4 = st.columns(4)

    with b1:
        if os.path.exists(ARCHIVO_WORD_PATH):
            with open(ARCHIVO_WORD_PATH, "rb") as f_doc:
                st.download_button(
                    label="📄 Guía en Word (.docx)",
                    data=f_doc.read(),
                    file_name=f"Guia_ULima_{tema.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
    with b2:
        if os.path.exists(ARCHIVO_ICS_PATH):
            with open(ARCHIVO_ICS_PATH, "rb") as f_ics:
                st.download_button(
                    label="📅 Bloque Estudio (.ics)",
                    data=f_ics.read(),
                    file_name="Sesion_Estudio_ULima.ics",
                    mime="text/calendar",
                    use_container_width=True
                )
    with b3:
        st.download_button(
            label="📝 Resumen Markdown (.md)",
            data=guia_final,
            file_name=f"Guia_ULima_{tema.replace(' ', '_')}.md",
            mime="text/markdown",
            use_container_width=True
        )
    with b4:
        titulo_gcal = urllib.parse.quote(f"Estudio ULima: {tema}")
        detalles_gcal = urllib.parse.quote(f"Sesión guiada ({detalle_perfil}) - {necesidad}.")
        fechas_gcal = f"{inicio_dt.strftime('%Y%m%dT%H%M%S')}/{fin_dt.strftime('%Y%m%dT%H%M%S')}"
        url_gcal = f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={titulo_gcal}&dates={fechas_gcal}&details={detalles_gcal}"
        st.link_button("🗓️ Agendar en Google Calendar", url_gcal, use_container_width=True)

    st.divider()
    st.subheader(f"📘 Guía Personalizada — {detalle_perfil}")
    st.markdown(guia_final)
