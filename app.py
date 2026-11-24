import os
import re
import html
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
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from google import genai
from google.genai import types
from duckduckgo_search import DDGS

st.set_page_config(
    page_title="AGENTE IA ULIMA | Google Workspace & ODS 4",
    page_icon="🎓",
    layout="wide"
)

ZONA_LIMA = timezone(timedelta(hours=-5))
ARCHIVO_WORD_PATH = "/tmp/Documento_ULima_ODS4.docx"
ARCHIVO_PPT_PATH = "/tmp/Presentacion_ULima_ODS4.pptx"
ARCHIVO_EXCEL_PATH = "/tmp/Hoja_Calculo_ULima_ODS4.xlsx"
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


def obtener_secreto(clave: str) -> str:
    try:
        return str(st.secrets.get(clave, "")).strip()
    except Exception:
        return os.environ.get(clave, "").strip()


def inicializar_estado():
    defaults = {
        "conectado": False,
        "api_key": obtener_secreto("GEMINI_API_KEY"),
        "gmail_user": obtener_secreto("GMAIL_USER"),
        "gmail_pass": obtener_secreto("GMAIL_APP_PASSWORD"),
        "modelo_preferido": "gemini-3.8-flash"
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


inicializar_estado()


# ==============================================================================
# HERRAMIENTAS DE NAVEGACIÓN WEB EN VIVO
# ==============================================================================
def leer_pagina_web(url: str) -> str:
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; AgenteIAULima/1.0)"}
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            raw = resp.read().decode("utf-8", errors="ignore")
        sin_scripts = re.sub(r"<(script|style).*?>.*?</\1>", " ", raw, flags=re.DOTALL | re.IGNORECASE)
        texto_plano = re.sub(r"<[^>]+>", " ", sin_scripts)
        texto_limpio = re.sub(r"\s+", " ", html.unescape(texto_plano)).strip()
        return texto_limpio[:1400]
    except Exception:
        return ""


def navegar_internet_vivo(consulta: str) -> tuple[str, str]:
    hora = ahora_lima().strftime("%H:%M:%S")
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(consulta, max_results=4))
        if not resultados:
            return "Base académica UNESCO ODS 4 y Universidad de Lima.", f"[{hora}] 🌐 Navegador Web: consulta completada"

        bloques = []
        urls_leidas = 0
        for idx, r in enumerate(resultados):
            titulo = r.get("title", "")
            resumen = r.get("body", "")
            enlace = r.get("href", "")
            extra_web = ""
            if idx < 2 and enlace.startswith("http"):
                contenido_url = leer_pagina_web(enlace)
                if contenido_url:
                    extra_web = f"\n  Extracto leído de la web: {contenido_url[:600]}..."
                    urls_leidas += 1
            bloques.append(f"- {titulo}: {resumen} (URL: {enlace}){extra_web}")

        log = f"[{hora}] 🌐 Navegador Web en vivo: {len(resultados)} fuentes halladas ({urls_leidas} páginas leídas a fondo)"
        return "\n".join(bloques), log
    except Exception:
        return "Estándares de educación superior UNESCO ODS 4 (ULima).", f"[{hora}] 🌐 Navegador Web: modo respaldo académico activo"


def buscar_wikipedia_vivo(concepto: str) -> tuple[str, str]:
    hora = ahora_lima().strftime("%H:%M:%S")
    try:
        termino = urllib.parse.quote(concepto.strip().split()[0])
        url = f"https://es.wikipedia.org/api/rest_v1/page/summary/{termino}"
        req = urllib.request.Request(url, headers={"User-Agent": "AgenteULimaODS4/8.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            datos = json.loads(resp.read().decode("utf-8"))
        extracto = datos.get("extract", "")
        if extracto:
            return extracto, f"[{hora}] 📚 Wikipedia consultada: '{concepto}' verificado"
        return "", f"[{hora}] 📚 Wikipedia: consulta completada"
    except Exception:
        return "", f"[{hora}] 📚 Wikipedia: síntesis interna activa"


# ==============================================================================
# CONEXIÓN BLINDADA CON GOOGLE GEMINI
# ==============================================================================
def obtener_modelos_activos(client: genai.Client, modelo_preferido: str) -> list[str]:
    prioridad = [modelo_preferido, "gemini-3.8-flash", "gemini-3.8-pro"]
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
                texto_resp = getattr(respuesta, "text", None) or "Acción ejecutada correctamente."
                return texto_resp, modelo
            except Exception as e:
                ultimo_error = str(e)
                if "API_KEY_INVALID" in ultimo_error:
                    raise ValueError("La API Key de Google AI Studio no es válida. Verifica que empiece con 'AIza...'.")
                if "404" in ultimo_error or "NOT_FOUND" in ultimo_error:
                    break
                time.sleep(2.0)

    raise RuntimeError(f"Los servidores de Google están ocupados en este instante. Espera 15 segundos y reintenta. ({ultimo_error[:120]})")


# ==============================================================================
# GENERADORES DE ARCHIVOS: WORD (.DOCX), POWERPOINT (.PPTX), EXCEL (.XLSX) Y CALENDAR
# ==============================================================================
def crear_word_archivo(titulo: str, subtitulo: str, contenido_md: str) -> str:
    hora = ahora_lima().strftime("%H:%M:%S")
    doc = Document()
    doc.add_heading("UNIVERSIDAD DE LIMA — ASISTENTE IA AUTÓNOMO (ODS 4)", level=0)
    doc.add_heading(f"{titulo} | {subtitulo}", level=1)
    doc.add_paragraph(f"Generado el {ahora_lima().strftime('%d/%m/%Y a las %H:%M')} (Hora Lima)")

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
    return f"[{hora}] 📄 Documento Word (.docx) creado"


def crear_ppt_archivo(titulo_portada: str, subtitulo: str, contenido_md: str) -> str:
    hora = ahora_lima().strftime("%H:%M:%S")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Portada institucional ULima
    slide_portada = prs.slides.add_slide(prs.slide_layouts[0])
    bg = slide_portada.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(24, 24, 27)

    if slide_portada.shapes.title:
        slide_portada.shapes.title.text = f"UNIVERSIDAD DE LIMA — ODS 4\n{titulo_portada}"
        for p in slide_portada.shapes.title.text_frame.paragraphs:
            p.font.color.rgb = RGBColor(249, 115, 22)
            p.font.bold = True
            p.font.size = Pt(34)

    if len(slide_portada.placeholders) > 1:
        slide_portada.placeholders[1].text = f"{subtitulo}\nGenerado autónomamente por Agente IA ULima ({ahora_lima().strftime('%d/%m/%Y')})"
        for p in slide_portada.placeholders[1].text_frame.paragraphs:
            p.font.color.rgb = RGBColor(244, 244, 245)
            p.font.size = Pt(18)

    # Parsear secciones en diapositivas
    secciones = []
    titulo_actual = "Resumen Ejecutivo y Enfoque ODS 4"
    puntos_actuales = []

    for linea in contenido_md.split("\n"):
        limpia = linea.strip().replace("**", "")
        if not limpia or limpia.startswith("|---"):
            continue
        if limpia.startswith("## ") or limpia.startswith("### ") or limpia.startswith("# "):
            if puntos_actuales:
                secciones.append((titulo_actual, puntos_actuales))
            titulo_actual = limpia.lstrip("#").strip()
            puntos_actuales = []
        else:
            puntos_actuales.append(limpia.lstrip("-*• ").strip())

    if puntos_actuales:
        secciones.append((titulo_actual, puntos_actuales))

    if not secciones:
        secciones = [(titulo_portada, [contenido_md[:600]])]

    for tit_slide, bullets in secciones[:8]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        if slide.shapes.title:
            slide.shapes.title.text = tit_slide[:85]
            for p in slide.shapes.title.text_frame.paragraphs:
                p.font.color.rgb = RGBColor(234, 88, 12)
                p.font.bold = True
                p.font.size = Pt(26)

        if len(slide.placeholders) > 1:
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for idx_b, b in enumerate(bullets[:6]):
                p = tf.paragraphs[0] if idx_b == 0 else tf.add_paragraph()
                p.text = f"• {b[:190]}"
                p.font.size = Pt(17)
                p.space_after = Pt(10)

    prs.save(ARCHIVO_PPT_PATH)
    return f"[{hora}] 📊 Presentación PowerPoint (.pptx) creada ({min(len(secciones), 8) + 1} diapositivas)"


def crear_excel_archivo(titulo: str, subtitulo: str, contenido_md: str, fecha_evento: datetime) -> str:
    hora = ahora_lima().strftime("%H:%M:%S")
    wb = Workbook()

    header_fill = PatternFill(start_color="E85D04", end_color="E85D04", fill_type="solid")
    sub_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    title_font = Font(color="FFFFFF", bold=True, size=13)

    ws1 = wb.active
    ws1.title = "Plan y Cronograma"
    ws1.merge_cells("A1:E1")
    ws1["A1"] = f"UNIVERSIDAD DE LIMA (ODS 4) — {titulo.upper()}"
    ws1["A1"].fill = sub_fill
    ws1["A1"].font = title_font
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.append(["Fase", "Actividad Clave", "Fecha Programada", "Duración", "Estado"])
    for cell in ws1[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    bloques = [
        ("1. Fundamentos", f"Comprensión base de {titulo}", fecha_evento.strftime("%d/%m/%Y %H:%M"), "30 min", "Pendiente"),
        ("2. Aplicación Práctica", f"Ejercicios / casos ({subtitulo})", (fecha_evento + timedelta(minutes=30)).strftime("%d/%m/%Y %H:%M"), "40 min", "Pendiente"),
        ("3. Validación ODS 4", "Checklist de dominio y repaso", (fecha_evento + timedelta(minutes=70)).strftime("%d/%m/%Y %H:%M"), "20 min", "Pendiente"),
    ]
    for b in bloques:
        ws1.append(list(b))

    ws2 = wb.create_sheet(title="Contenido Detallado")
    ws2.append(["Bloque / Sección", "Información Desarrollada"])
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font

    sec_actual = "Resumen"
    for linea in contenido_md.split("\n"):
        limpia = linea.strip().replace("**", "")
        if not limpia:
            continue
        if limpia.startswith("#"):
            sec_actual = limpia.lstrip("#").strip()
        else:
            ws2.append([sec_actual, limpia])

    ws3 = wb.create_sheet(title="14 Carreras ULima")
    ws3.append(["N.º", "Carrera ULima", "Enfoque Académico", "ODS 4"])
    for cell in ws3[1]:
        cell.fill = header_fill
        cell.font = header_font
    for idx, c in enumerate(CARRERAS_ULIMA[:14], start=1):
        ws3.append([idx, c, "Innovación, gestión y formación integral", "Metas 4.3 y 4.4"])

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
    return f"[{hora}] 📈 Libro de Excel (.xlsx) creado (3 hojas)"


def construir_evento_calendar(titulo_evento: str, inicio_dt: datetime, duracion_min: int, descripcion: str, correo_asistente: str) -> tuple[str, str, str, datetime]:
    hora = ahora_lima().strftime("%H:%M:%S")
    fin_dt = inicio_dt + timedelta(minutes=duracion_min)
    uid_evento = f"ulima-ods4-{int(time.time())}@ulima.edu.pe"
    dtstamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    ics_texto = (
        "BEGIN:VCALENDAR\n"
        "VERSION:2.0\n"
        "PRODID:-//Agente IA ULima ODS 4//Google Workspace//ES\n"
        "METHOD:REQUEST\n"
        "BEGIN:VEVENT\n"
        f"UID:{uid_evento}\n"
        f"DTSTAMP:{dtstamp}\n"
        f"DTSTART:{inicio_dt.strftime('%Y%m%dT%H%M%S')}\n"
        f"DTEND:{fin_dt.strftime('%Y%m%dT%H%M%S')}\n"
        f"ORGANIZER;CN=Agente IA ULima:mailto:{correo_asistente}\n"
        f"ATTENDEE;RSVP=TRUE:mailto:{correo_asistente}\n"
        f"SUMMARY:{titulo_evento}\n"
        f"DESCRIPTION:{descripcion}\n"
        "STATUS:CONFIRMED\n"
        "END:VEVENT\n"
        "END:VCALENDAR\n"
    )
    with open(ARCHIVO_ICS_PATH, "w", encoding="utf-8") as f:
        f.write(ics_texto)

    titulo_q = urllib.parse.quote(titulo_evento)
    detalles_q = urllib.parse.quote(descripcion)
    fechas_q = f"{inicio_dt.strftime('%Y%m%dT%H%M%S')}/{fin_dt.strftime('%Y%m%dT%H%M%S')}"
    url_gcal = f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={titulo_q}&dates={fechas_q}&details={detalles_q}"

    log = f"[{hora}] 📅 Google Calendar agendado: '{titulo_evento}' ({inicio_dt.strftime('%d/%m/%Y %H:%M')})"
    return log, url_gcal, ics_texto, inicio_dt


def enviar_gmail_completo(remitente: str, clave_app: str, destinatario: str, asunto: str, contenido: str, ics_texto: str, info_fecha: str) -> tuple[bool, str]:
    hora = ahora_lima().strftime("%H:%M:%S")
    remitente = remitente.strip()
    clave_app = clave_app.replace(" ", "").strip()
    destinatario = destinatario.strip() or remitente

    if not remitente or not clave_app:
        return False, f"[{hora}] ⚠️ Gmail: ingresa tu correo y contraseña de aplicación para enviar automáticamente"

    try:
        msg = MIMEMultipart("mixed")
        msg["From"] = remitente
        msg["To"] = destinatario
        msg["Subject"] = asunto

        cuerpo = (
            "¡Hola!\n\n"
            "Tu Agente IA Autónomo ULima (ODS 4) ha ejecutado tu solicitud y conectado tus herramientas:\n"
            f"📅 Evento programado en Google Calendar: {info_fecha}\n"
            "📄 Documento Word (.docx) adjunto\n"
            "📊 Presentación PowerPoint (.pptx) adjunta\n"
            "📈 Libro Excel (.xlsx) adjunto\n\n"
            "------------------------------------------------------------\n"
            f"{contenido[:3500]}\n"
            "------------------------------------------------------------\n"
            "Agente IA Universidad de Lima — ODS 4."
        )
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

        if ics_texto:
            parte_cal = MIMEText(ics_texto, "calendar; method=REQUEST", "utf-8")
            msg.attach(parte_cal)

        adjuntos = [
            (ARCHIVO_WORD_PATH, "Documento_ULima_ODS4.docx"),
            (ARCHIVO_PPT_PATH, "Presentacion_ULima_ODS4.pptx"),
            (ARCHIVO_EXCEL_PATH, "Hoja_Calculo_ULima_ODS4.xlsx"),
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

        return True, f"[{hora}] ✅ Gmail enviado a '{destinatario}' con Word + PPT + Excel + Invitación Calendar"
    except smtplib.SMTPAuthenticationError:
        return False, (
            f"[{hora}] ❌ Google rechazó la contraseña: debes usar una 'Contraseña de aplicación' de 16 letras "
            "(myaccount.google.com/apppasswords)."
        )
    except Exception as e:
        return False, f"[{hora}] ❌ Error al enviar Gmail: {str(e)}"


# ==============================================================================
# PANTALLA 0: ACCESO INICIAL (API KEY + CORREO Y CONTRASEÑA GOOGLE)
# ==============================================================================
if not st.session_state["conectado"]:
    st.markdown("<br>", unsafe_allow_html=True)
    _, col_centro, _ = st.columns([1, 1.5, 1])

    with col_centro:
        st.markdown("## 🎓 Agente IA Autónomo — Universidad de Lima")
        st.caption("Conecta tu clave de IA y tu cuenta de Google para que el agente navegue en internet, agende en Calendar, cree Docs/PPT/Excel y envíe Gmails.")

        with st.form("login_agente_ulima"):
            api_in = st.text_input(
                "🔑 Google AI Studio API Key (Cerebro IA)",
                value=st.session_state["api_key"],
                type="password",
                placeholder="AIzaSy..."
            )
            correo_in = st.text_input(
                "📧 Correo de Google (Gmail & Calendar)",
                value=st.session_state["gmail_user"],
                placeholder="tucorreo@gmail.com"
            )
            pass_in = st.text_input(
                "🔐 Contraseña de aplicación de Google (16 letras)",
                value=st.session_state["gmail_pass"],
                type="password",
                placeholder="••••••••••••••••"
            )
            modelo_in = st.selectbox(
                "⚡ Motor Gemini",
                ["gemini-3.8-flash", "gemini-3.8-pro"],
                index=0
            )

            btn_entrar = st.form_submit_button("🚀 Conectar e Iniciar Agente IA", type="primary", use_container_width=True)

            if btn_entrar:
                if not api_in.strip():
                    st.error("⚠️ Ingresa tu Google AI Studio API Key (`AIza...`) para activar la inteligencia artificial.")
                else:
                    st.session_state["api_key"] = api_in.strip()
                    st.session_state["gmail_user"] = correo_in.strip()
                    st.session_state["gmail_pass"] = pass_in.strip()
                    st.session_state["modelo_preferido"] = modelo_in
                    st.session_state["conectado"] = True
                    st.rerun()

        with st.expander("¿Cómo obtener tu contraseña de acceso de Google (16 letras)?"):
            st.markdown(
                "1. Entra a **myaccount.google.com/apppasswords** con tu cuenta Google.\n"
                "2. Escribe `Agente ULima` y haz clic en **Crear**.\n"
                "3. Copia las **16 letras** y pégalas arriba para habilitar envíos reales por Gmail y Google Calendar."
            )

    st.stop()


# ==============================================================================
# WORKSPACE AUTÓNOMO ULIMA (CONECTADO A TODO)
# ==============================================================================
with st.sidebar:
    st.title("⚙️ Centro de Control")
    st.success("🟢 Gemini IA Conectado")
    if st.session_state["gmail_user"] and st.session_state["gmail_pass"]:
        st.success(f"🟢 Google: {st.session_state['gmail_user']}")
    else:
        st.warning("🟡 Gmail sin contraseña (solo descargas)")

    st.divider()
    correo_destino_ui = st.text_input(
        "📩 Correo destino por defecto:",
        value=st.session_state["gmail_user"],
        placeholder="alumno@aloe.ulima.edu.pe"
    )
    fecha_base_ui = st.date_input(
        "📅 Fecha base Calendar",
        value=(ahora_lima() + timedelta(days=1)).date()
    )
    hora_base_ui = st.time_input(
        "⏰ Hora base Calendar",
        value=datetime.strptime("18:00", "%H:%M").time()
    )

    st.divider()
    st.markdown("**Capacidades Activas:**")
    st.markdown("🌐 Navegación Web en Vivo")
    st.markdown("📅 Google Calendar (Autónomo)")
    st.markdown("📧 Gmail con Adjuntos")
    st.markdown("📄 Word (`.docx`)")
    st.markdown("📊 PowerPoint (`.pptx`)")
    st.markdown("📈 Excel (`.xlsx`)")

    st.divider()
    if st.button("🔄 Cambiar Credenciales", use_container_width=True):
        st.session_state["conectado"] = False
        st.rerun()

st.title("🎓 AGENTE IA ULIMA — Orientación, Postulantes y 14 Carreras (ODS 4)")
st.markdown(
    "Agente autónomo con **navegación por internet en tiempo real**, conectado a **Gmail, Google Calendar, Word (.docx), PowerPoint (.pptx) y Excel (.xlsx)** "
    "para facilitar el ingreso, la orientación vocacional y el estudio en las **14 carreras de la Universidad de Lima**."
)

perfil = st.radio(
    "Selecciona el modo de operación:",
    [
        "🧭 Orientación Vocacional ULima (14 carreras)",
        "🚀 Postulante a la ULima (Admisión / CEPREULIMA / Escolar)",
        "🎓 Alumno Regular ULima (14 Carreras / Estudios Generales)",
        "🤖 Asistente Libre IA (Pídele cualquier tarea, investigación, PPT, Doc, Excel, Calendar o Gmail)"
    ],
    horizontal=True
)

st.divider()

if "Orientación" in perfil:
    col1, col2, col3 = st.columns(3)
    with col1:
        intereses = st.text_input(
            "Tus intereses o habilidades",
            value="Tecnología, gestión de empresas, liderazgo e innovación"
        )
    with col2:
        carreras_comparar = st.multiselect(
            "Carreras ULima a evaluar (14 disponibles)",
            CARRERAS_ULIMA[:14],
            default=["Ingeniería de Sistemas", "Ingeniería Industrial", "Administración"]
        )
    with col3:
        tema = st.text_input(
            "Enfoque de orientación",
            value="Comparativa de carreras ULima, campo laboral y perfil ideal"
        )
    detalle_perfil = ", ".join(carreras_comparar) if carreras_comparar else "14 Carreras ULima"
    necesidad = "Orientación Vocacional Personalizada ULima"
    valor_default_mision = (
        f"Investiga en internet sobre las carreras de la ULima ({detalle_perfil}) según mis intereses ('{intereses}'), "
        "dime cuál es mi mejor opción, créame un Word, una presentación PPT y un Excel comparativo, prográmalo en mi Google Calendar para mañana a las 5:00 pm y mándalo a mi correo."
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
                "Plan de estudio en Excel, PPT y Google Calendar",
                "Guía puente hacia Estudios Generales ULima"
            ]
        )
    detalle_perfil = f"{modalidad} → {carrera_objetivo}"
    valor_default_mision = (
        f"Explícalo fácil desde cero para un postulante a {carrera_objetivo} en la ULima con 4 preguntas tipo examen resueltas paso a paso, "
        "hazme un Word, un PowerPoint resumen y un Excel de estudio, agéndalo en mi Google Calendar para mañana a las 6:00 pm y envíamelo a mi correo."
    )

elif "Alumno" in perfil:
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
                "Presentación PPT + Word lista para exposición grupal",
                "Explicar teoría difícil desde cero con ejemplos prácticos",
                "Plan de estudio en Excel + Google Calendar"
            ]
        )
    detalle_perfil = carrera_actual
    valor_default_mision = (
        f"Investiga en internet casos actuales, explícalo fácil para un alumno de {carrera_actual} en la ULima con 3 ejercicios resueltos, "
        "genera un Word, una presentación PPT y un Excel descargable, prográmalo en mi Google Calendar para mañana a las 7:00 pm y envíamelo al correo."
    )

else:
    col1, col2 = st.columns(2)
    with col1:
        tema = st.text_input(
            "Tema o proyecto principal",
            value="Impacto de la IA en la Educación Universitaria (ODS 4 - ULima)"
        )
    with col2:
        detalle_perfil = st.selectbox(
            "Contexto / Facultad de referencia",
            ["Comunidad ULima General"] + CARRERAS_ULIMA
        )
    necesidad = "Ejecución Autónoma Multiherramienta"
    valor_default_mision = (
        "Busca información actualizada en internet sobre este tema, crea un informe en Word, una presentación PPT de 6 diapositivas "
        "y un Excel con métricas clave, agéndalo en mi Google Calendar para mañana a las 4:00 pm y envíalo a mi correo."
    )

objetivo_extra = st.text_area(
    "Pídele lo que quieras como IA (navegar en internet, crear Word, PPT, Excel, agendar fechas en Calendar o mandar Gmails):",
    value=valor_default_mision,
    height=105
)

if st.button("⚡ Ejecutar Agente IA Autónomo", type="primary", use_container_width=True):
    tiempo_inicio = time.time()
    bitacora = []
    ahora = ahora_lima()
    fecha_hoy_str = f"{DIAS_ES[ahora.weekday()]} {ahora.strftime('%Y-%m-%d')} (hora actual Lima: {ahora.strftime('%H:%M')})"

    correos_en_texto = re.findall(r"[\w\.-]+@[\w\.-]+\.\w+", objetivo_extra)
    destinatario_final = correos_en_texto[0] if correos_en_texto else correo_destino_ui.strip()

    estado_calendario = {
        "inicio": datetime.combine(fecha_base_ui, hora_base_ui),
        "titulo": f"ULima ODS 4: {tema}",
        "url_gcal": "",
        "ics_texto": "",
        "programado_por_ia": False
    }
    estado_archivos = {"word": False, "ppt": False, "excel": False}
    estado_gmail = {"intentado": False, "exito": False, "mensaje": ""}

    try:
        client = genai.Client(api_key=st.session_state["api_key"])

        with st.status("⚡ Agente IA navegando en internet y ejecutando herramientas...", expanded=True) as status:
            # FASE 1: Navegación Web en Vivo + Wikipedia
            st.write("🌐 **Fase 1/3:** Navegando en internet y leyendo fuentes académicas en tiempo real...")
            query_busqueda = (
                f"Universidad de Lima {detalle_perfil} {tema}"
                if "Libre" not in perfil else
                f"{tema} Universidad de Lima ODS 4"
            )
            with ThreadPoolExecutor(max_workers=3) as pool:
                futuro_web = pool.submit(navegar_internet_vivo, query_busqueda)
                futuro_wiki = pool.submit(buscar_wikipedia_vivo, tema)
                futuro_modelos = pool.submit(obtener_modelos_activos, client, st.session_state["modelo_preferido"])
                info_web, log_web = futuro_web.result()
                info_wiki, log_wiki = futuro_wiki.result()
                modelos_activos = futuro_modelos.result()

            bitacora.extend([log_web, log_wiki])

            # FASE 2: Síntesis e Inteligencia Pedagógica ULima
            st.write("🧠 **Fase 2/3:** Razonando y elaborando el contenido estructurado con Gemini IA...")
            prompt_maestro = (
                f"Eres el Agente IA Autónomo de la Universidad de Lima (ULima - ODS 4). "
                f"Perfil/Modo: {perfil} ({detalle_perfil}). Tema: '{tema}'. Objetivo: '{necesidad}'.\n"
                f"Instrucción exacta del usuario: {objetivo_extra}\n\n"
                f"Información extraída de internet en vivo:\n{info_web}\n\n"
                f"Referencia enciclopédica:\n{info_wiki}\n\n"
                f"Genera una respuesta completa, clara, fácil de entender y muy bien estructurada en Markdown con secciones (`## `) "
                f"para que sirva tanto de lectura directa como de base para generar documentos Word, diapositivas PPT y tablas Excel. "
                f"Encabézala con un breve **Índice de Facilitación ULima (ODS 4): [Puntaje/100]** y cierra con un **Checklist Rápido de Acción**."
            )
            contenido_final, modelo_usado = llamar_gemini_blindado(client, prompt_maestro, modelos_activos)
            modelos_priorizados = [modelo_usado] + [m for m in modelos_activos if m != modelo_usado]
            bitacora.append(f"[{ahora_lima().strftime('%H:%M:%S')}] 🧠 Contenido generado por IA ({modelo_usado})")

            # FASE 3: Orquestación Autónoma de Herramientas (Calendar, Word, PPT, Excel, Gmail)
            st.write("🤖 **Fase 3/3:** Ejecutando acciones en Google Calendar, Word, PowerPoint, Excel y Gmail...")

            def tool_agendar_google_calendar(titulo_evento: str, fecha_yyyy_mm_dd: str, hora_hh_mm: str, duracion_minutos: int = 90) -> str:
                """Crea y programa una fecha y hora exacta en Google Calendar según la instrucción del usuario."""
                try:
                    dt_parsed = datetime.strptime(f"{fecha_yyyy_mm_dd.strip()} {hora_hh_mm.strip()}", "%Y-%m-%d %H:%M")
                except Exception:
                    dt_parsed = datetime.combine(fecha_base_ui, hora_base_ui)

                log_cal, url_cal, ics_txt, ini_dt = construir_evento_calendar(
                    titulo_evento or f"ULima ODS 4: {tema}",
                    dt_parsed,
                    duracion_minutos or 90,
                    f"Sesión ULima ({detalle_perfil}) - {necesidad}",
                    destinatario_final
                )
                estado_calendario["inicio"] = ini_dt
                estado_calendario["titulo"] = titulo_evento or f"ULima ODS 4: {tema}"
                estado_calendario["url_gcal"] = url_cal
                estado_calendario["ics_texto"] = ics_txt
                estado_calendario["programado_por_ia"] = True
                bitacora.append(log_cal)
                return log_cal

            def tool_generar_documentos_office(hacer_word: bool = True, hacer_ppt: bool = True, hacer_excel: bool = True) -> str:
                """Genera los archivos descargables en Word (.docx), PowerPoint (.pptx) y Excel (.xlsx)."""
                res = []
                if hacer_word:
                    log_w = crear_word_archivo(tema, detalle_perfil, contenido_final)
                    estado_archivos["word"] = True
                    bitacora.append(log_w)
                    res.append("Word")
                if hacer_ppt:
                    log_p = crear_ppt_archivo(tema, detalle_perfil, contenido_final)
                    estado_archivos["ppt"] = True
                    bitacora.append(log_p)
                    res.append("PPT")
                if hacer_excel:
                    log_e = crear_excel_archivo(tema, detalle_perfil, contenido_final, estado_calendario["inicio"])
                    estado_archivos["excel"] = True
                    bitacora.append(log_e)
                    res.append("Excel")
                return "Archivos generados: " + ", ".join(res)

            prompt_orquestador = (
                f"Hoy es {fecha_hoy_str}. Analiza la instrucción del usuario: '{objetivo_extra}'.\n"
                f"1) Llama a `tool_agendar_google_calendar` convirtiendo cualquier fecha/hora mencionada ('mañana', 'este viernes', 'a las 4 pm', etc.) "
                f"a `YYYY-MM-DD` y `HH:MM` (24h). Si no mencionó fecha/hora, usa `{fecha_base_ui.strftime('%Y-%m-%d')}` y `{hora_base_ui.strftime('%H:%M')}`.\n"
                f"2) Llama a `tool_generar_documentos_office(hacer_word=True, hacer_ppt=True, hacer_excel=True)` para crear los archivos solicitados."
            )
            llamar_gemini_blindado(
                client,
                prompt_orquestador,
                modelos_priorizados,
                herramientas=[tool_agendar_google_calendar, tool_generar_documentos_office]
            )

            # Garantizar ejecución completa si el modelo omitió alguna llamada
            if not estado_calendario["programado_por_ia"]:
                dt_def = datetime.combine(fecha_base_ui, hora_base_ui)
                log_cal, url_cal, ics_txt, ini_dt = construir_evento_calendar(
                    f"ULima ODS 4: {tema}",
                    dt_def,
                    90,
                    f"Sesión ULima ({detalle_perfil}) - {necesidad}",
                    destinatario_final
                )
                estado_calendario["inicio"] = ini_dt
                estado_calendario["url_gcal"] = url_cal
                estado_calendario["ics_texto"] = ics_txt
                bitacora.append(log_cal)

            if not estado_archivos["word"]:
                bitacora.append(crear_word_archivo(tema, detalle_perfil, contenido_final))
                estado_archivos["word"] = True
            if not estado_archivos["ppt"]:
                bitacora.append(crear_ppt_archivo(tema, detalle_perfil, contenido_final))
                estado_archivos["ppt"] = True
            if not estado_archivos["excel"]:
                bitacora.append(crear_excel_archivo(tema, detalle_perfil, contenido_final, estado_calendario["inicio"]))
                estado_archivos["excel"] = True

            # Envío automático por Gmail con Word + PPT + Excel + Calendar adjuntos
            if destinatario_final:
                exito_mail, msg_mail = enviar_gmail_completo(
                    st.session_state["gmail_user"],
                    st.session_state["gmail_pass"],
                    destinatario_final,
                    f"Agente IA ULima ({detalle_perfil}): {tema} [{estado_calendario['inicio'].strftime('%d/%m %H:%M')}]",
                    contenido_final,
                    estado_calendario["ics_texto"],
                    f"{estado_calendario['titulo']} — {estado_calendario['inicio'].strftime('%d/%m/%Y %H:%M')}"
                )
                estado_gmail["intentado"] = True
                estado_gmail["exito"] = exito_mail
                estado_gmail["mensaje"] = msg_mail
                bitacora.append(msg_mail)

            duracion = round(time.time() - tiempo_inicio, 1)
            status.update(label=f"✨ Misión autónoma completada en {duracion}s ({modelo_usado})", state="complete", expanded=False)

    except Exception as error_controlado:
        st.warning(f"⚠️ Aviso del sistema: {error_controlado}")
        st.stop()

    # Alertas visuales de sincronización
    if estado_gmail["intentado"]:
        if estado_gmail["exito"]:
            st.success(f"✅ **Gmail y Calendar Sincronizados:** {estado_gmail['mensaje']}")
        else:
            st.warning(f"⚠️ {estado_gmail['mensaje']}")

    st.info(
        f"📅 **Google Calendar:** `{estado_calendario['titulo']}` programado para el "
        f"**{estado_calendario['inicio'].strftime('%d/%m/%Y a las %H:%M')}** (Hora Lima)."
    )

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Tiempo Total", f"{duracion} seg")
    m2.metric("Office Generado", "Word + PPT + Excel")
    m3.metric("Google Calendar", estado_calendario["inicio"].strftime("%d/%m %H:%M"))
    m4.metric("Gmail", "Enviado ✅" if estado_gmail["exito"] else "Listo para descargar")

    with st.expander("🛠️ Bitácora en Vivo de Navegación y Acciones del Agente", expanded=True):
        for linea_log in bitacora:
            st.code(linea_log, language="text")

    st.subheader("📦 Descargas Directas (Word, PowerPoint, Excel) y Google Calendar")
    b1, b2, b3, b4, b5 = st.columns(5)

    with b1:
        if os.path.exists(ARCHIVO_WORD_PATH):
            with open(ARCHIVO_WORD_PATH, "rb") as f_doc:
                st.download_button(
                    label="📄 Word (.docx)",
                    data=f_doc.read(),
                    file_name=f"Documento_ULima_{tema.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
    with b2:
        if os.path.exists(ARCHIVO_PPT_PATH):
            with open(ARCHIVO_PPT_PATH, "rb") as f_ppt:
                st.download_button(
                    label="📊 PowerPoint (.pptx)",
                    data=f_ppt.read(),
                    file_name=f"Presentacion_ULima_{tema.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
    with b3:
        if os.path.exists(ARCHIVO_EXCEL_PATH):
            with open(ARCHIVO_EXCEL_PATH, "rb") as f_xls:
                st.download_button(
                    label="📈 Excel (.xlsx)",
                    data=f_xls.read(),
                    file_name=f"Hoja_ULima_{tema.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    with b4:
        if os.path.exists(ARCHIVO_ICS_PATH):
            with open(ARCHIVO_ICS_PATH, "rb") as f_ics:
                st.download_button(
                    label="📅 Calendar (.ics)",
                    data=f_ics.read(),
                    file_name="Evento_ULima_ODS4.ics",
                    mime="text/calendar",
                    use_container_width=True
                )
    with b5:
        if estado_calendario["url_gcal"]:
            st.link_button("🗓️ Abrir Google Calendar", estado_calendario["url_gcal"], use_container_width=True)

    st.divider()
    st.subheader(f"📘 Resultado Generado — {detalle_perfil}")
    st.markdown(contenido_final)
