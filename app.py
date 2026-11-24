import os
import re
import html
import time
import json
import base64
import urllib.request
import urllib.parse
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
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from duckduckgo_search import DDGS

os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"] = "1"

st.set_page_config(
    page_title="Agente IA ULima | Google Workspace & Web",
    page_icon="🎓",
    layout="wide"
)

ZONA_LIMA = timezone(timedelta(hours=-5))
DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]

SCOPES_GOOGLE = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/drive.file"
]

CARRERAS_ULIMA = [
    "Administración", "Arquitectura", "Comunicación", "Contabilidad y Finanzas",
    "Derecho", "Economía", "Ingeniería Ambiental", "Ingeniería Civil",
    "Ingeniería de Sistemas", "Ingeniería Industrial", "Ingeniería Mecatrónica",
    "Marketing", "Negocios Internacionales", "Psicología"
]


def ahora_lima() -> datetime:
    return datetime.now(ZONA_LIMA)


def obtener_secreto(clave: str, default: str = "") -> str:
    try:
        val = st.secrets.get(clave, "")
        return str(val).strip() if val else default
    except Exception:
        return os.environ.get(clave, default).strip()


def inicializar_estado():
    defaults = {
        "api_key": obtener_secreto("GEMINI_API_KEY"),
        "client_id": obtener_secreto("GOOGLE_CLIENT_ID"),
        "client_secret": obtener_secreto("GOOGLE_CLIENT_SECRET"),
        "redirect_uri": obtener_secreto("REDIRECT_URI", "https://agente-autonomo-ods4.streamlit.app"),
        "google_creds": None,
        "google_email": "",
        "modo_sin_oauth": False,
        "mensajes": [],
        "acciones_turno": [],
        "entregables_turno": {}
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


inicializar_estado()


def crear_flujo_oauth() -> Flow:
    client_config = {
        "web": {
            "client_id": st.session_state["client_id"],
            "client_secret": st.session_state["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [st.session_state["redirect_uri"]]
        }
    }
    flow = Flow.from_client_config(
        client_config,
        scopes=SCOPES_GOOGLE,
        redirect_uri=st.session_state["redirect_uri"]
    )
    flow.autogenerate_code_verifier = False
    return flow


# Capturar el regreso automático desde Google ("Continuar con Google")
codigo_oauth = st.query_params.get("code")
if codigo_oauth and st.session_state["google_creds"] is None:
    if st.session_state["client_id"] and st.session_state["client_secret"]:
        try:
            flujo = crear_flujo_oauth()
            flujo.code_verifier = None
            flujo.fetch_token(code=codigo_oauth)
            creds = flujo.credentials
            st.session_state["google_creds"] = creds

            gmail_srv = build("gmail", "v1", credentials=creds)
            perfil = gmail_srv.users().getProfile(userId="me").execute()
            st.session_state["google_email"] = perfil.get("emailAddress", "Usuario Google")
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.query_params.clear()
            st.error(f"Error al conectar con Google: {e}")


# ==============================================================================
# HERRAMIENTAS AUTÓNOMAS DEL AGENTE IA
# ==============================================================================
def registrar_accion(texto: str):
    hora = ahora_lima().strftime("%H:%M:%S")
    st.session_state["acciones_turno"].append(f"[{hora}] {texto}")


def buscar_en_internet(consulta: str) -> str:
    """Busca información actualizada en internet y lee el contenido de las páginas más relevantes."""
    registrar_accion(f"🌐 Buscando en internet: '{consulta}'")
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(consulta, max_results=4))
        if not resultados:
            return "No se encontraron resultados nuevos en la web."

        salida = []
        for i, r in enumerate(resultados):
            titulo = r.get("title", "")
            snippet = r.get("body", "")
            url = r.get("href", "")
            extracto = ""
            if i < 2 and url.startswith("http"):
                try:
                    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req, timeout=4) as resp:
                        raw = resp.read().decode("utf-8", errors="ignore")
                    limpio = re.sub(r"<(script|style).*?>.*?</\1>", " ", raw, flags=re.DOTALL | re.IGNORECASE)
                    limpio = re.sub(r"<[^>]+>", " ", limpio)
                    limpio = re.sub(r"\s+", " ", html.unescape(limpio)).strip()
                    extracto = f"\nContenido leído: {limpio[:700]}"
                except Exception:
                    pass
            salida.append(f"Título: {titulo}\nURL: {url}\nResumen: {snippet}{extracto}")
        return "\n---\n".join(salida)
    except Exception as e:
        return f"Error al buscar en internet: {e}"


def leer_pagina_web(url: str) -> str:
    """Abre y lee el texto completo de una dirección URL específica de internet."""
    registrar_accion(f"🔗 Leyendo página web: {url}")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            raw = resp.read().decode("utf-8", errors="ignore")
        limpio = re.sub(r"<(script|style).*?>.*?</\1>", " ", raw, flags=re.DOTALL | re.IGNORECASE)
        limpio = re.sub(r"<[^>]+>", " ", limpio)
        limpio = re.sub(r"\s+", " ", html.unescape(limpio)).strip()
        return limpio[:3000] if limpio else "La página no devolvió texto legible."
    except Exception as e:
        return f"No se pudo abrir la URL: {e}"


def ver_eventos_google_calendar(dias_adelante: int = 7) -> str:
    """Consulta los próximos eventos agendados en el Google Calendar del usuario."""
    registrar_accion(f"📅 Consultando tu Google Calendar (próximos {dias_adelante} días)")
    creds = st.session_state.get("google_creds")
    if not creds:
        return "Google Calendar no está conectado en este modo. Pide al usuario iniciar sesión con Google."
    try:
        cal = build("calendar", "v3", credentials=creds)
        ahora_iso = ahora_lima().isoformat()
        fin_iso = (ahora_lima() + timedelta(days=dias_adelante)).isoformat()
        eventos = cal.events().list(
            calendarId="primary",
            timeMin=ahora_iso,
            timeMax=fin_iso,
            singleEvents=True,
            orderBy="startTime",
            maxResults=15
        ).execute().get("items", [])

        if not eventos:
            return "Tu calendario está completamente libre en ese rango de días."
        lista = []
        for ev in eventos:
            inicio = ev.get("start", {}).get("dateTime", ev.get("start", {}).get("date", ""))
            lista.append(f"- {ev.get('summary', 'Sin título')} | Inicio: {inicio}")
        return "\n".join(lista)
    except Exception as e:
        return f"Error al leer Google Calendar: {e}"


def agendar_evento_google_calendar(titulo: str, fecha_yyyy_mm_dd: str, hora_hh_mm: str, duracion_minutos: int = 60, descripcion: str = "") -> str:
    """Crea y guarda un evento real en el Google Calendar del usuario."""
    registrar_accion(f"🗓️ Agendando en Google Calendar: '{titulo}' ({fecha_yyyy_mm_dd} {hora_hh_mm})")
    creds = st.session_state.get("google_creds")
    try:
        dt_inicio = datetime.strptime(f"{fecha_yyyy_mm_dd.strip()} {hora_hh_mm.strip()}", "%Y-%m-%d %H:%M").replace(tzinfo=ZONA_LIMA)
    except Exception:
        dt_inicio = (ahora_lima() + timedelta(days=1)).replace(hour=18, minute=0, second=0, microsecond=0)
    dt_fin = dt_inicio + timedelta(minutes=duracion_minutos or 60)

    titulo_q = urllib.parse.quote(titulo)
    det_q = urllib.parse.quote(descripcion)
    fechas_q = f"{dt_inicio.strftime('%Y%m%dT%H%M%S')}/{dt_fin.strftime('%Y%m%dT%H%M%S')}"
    link_rapido = f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={titulo_q}&dates={fechas_q}&details={det_q}"

    if not creds:
        st.session_state["entregables_turno"]["calendar_link"] = link_rapido
        return f"Enlace directo generado para Google Calendar (modo sin OAuth): {link_rapido}"

    try:
        cal = build("calendar", "v3", credentials=creds)
        evento = {
            "summary": titulo,
            "description": descripcion or "Agendado por Agente IA ULima",
            "start": {"dateTime": dt_inicio.isoformat(), "timeZone": "America/Lima"},
            "end": {"dateTime": dt_fin.isoformat(), "timeZone": "America/Lima"},
            "reminders": {"useDefault": False, "overrides": [{"method": "popup", "minutes": 30}]}
        }
        creado = cal.events().insert(calendarId="primary", body=evento).execute()
        html_link = creado.get("htmlLink", link_rapido)
        st.session_state["entregables_turno"]["calendar_link"] = html_link
        return f"Evento creado exitosamente en Google Calendar para el {dt_inicio.strftime('%d/%m/%Y a las %H:%M')}. Enlace: {html_link}"
    except Exception as e:
        st.session_state["entregables_turno"]["calendar_link"] = link_rapido
        return f"Error al insertar en Calendar API ({e}), enlace alternativo: {link_rapido}"


def leer_ultimos_correos_gmail(cantidad: int = 5) -> str:
    """Lee los últimos correos recibidos en la bandeja de entrada de Gmail del usuario."""
    registrar_accion(f"📥 Leyendo los últimos {cantidad} correos de tu Gmail")
    creds = st.session_state.get("google_creds")
    if not creds:
        return "Gmail no está conectado en este modo. Pide al usuario iniciar sesión con Google."
    try:
        gmail = build("gmail", "v1", credentials=creds)
        mensajes = gmail.users().messages().list(userId="me", maxResults=min(cantidad, 10)).execute().get("messages", [])
        if not mensajes:
            return "No hay mensajes recientes en tu bandeja."
        resumenes = []
        for m in mensajes:
            detalle = gmail.users().messages().get(userId="me", id=m["id"], format="metadata", metadataHeaders=["From", "Subject", "Date"]).execute()
            headers = {h["name"]: h["value"] for h in detalle.get("payload", {}).get("headers", [])}
            snippet = detalle.get("snippet", "")
            resumenes.append(f"De: {headers.get('From', '')} | Asunto: {headers.get('Subject', '')} | Resumen: {snippet[:140]}")
        return "\n".join(resumenes)
    except Exception as e:
        return f"Error al leer Gmail: {e}"


def enviar_correo_gmail(destinatario: str, asunto: str, cuerpo: str, adjuntar_archivos: bool = True) -> str:
    """Envía un correo electrónico real desde el Gmail del usuario, adjuntando automáticamente Word, PPT o Excel si fueron creados."""
    dest = destinatario.strip() or st.session_state.get("google_email", "")
    registrar_accion(f"📧 Enviando correo vía Gmail a: {dest}")
    creds = st.session_state.get("google_creds")

    if not creds:
        url_mail = f"https://mail.google.com/mail/?view=cm&fs=1&to={urllib.parse.quote(dest)}&su={urllib.parse.quote(asunto)}&body={urllib.parse.quote(cuerpo[:1500])}"
        st.session_state["entregables_turno"]["gmail_link"] = url_mail
        return f"Gmail listo para enviar en 1 clic (modo sin OAuth): {url_mail}"

    try:
        gmail = build("gmail", "v1", credentials=creds)
        msg = MIMEMultipart()
        msg["To"] = dest
        msg["Subject"] = asunto

        links_extra = []
        ent = st.session_state.get("entregables_turno", {})
        if ent.get("docs_link"):
            links_extra.append(f"Google Docs: {ent['docs_link']}")
        if ent.get("slides_link"):
            links_extra.append(f"Google Slides: {ent['slides_link']}")
        if ent.get("sheets_link"):
            links_extra.append(f"Google Sheets: {ent['sheets_link']}")
        if ent.get("calendar_link"):
            links_extra.append(f"Google Calendar: {ent['calendar_link']}")

        texto_final = cuerpo
        if links_extra:
            texto_final += "\n\n--- Enlaces creados en tu Google Workspace ---\n" + "\n".join(links_extra)

        msg.attach(MIMEText(texto_final, "plain", "utf-8"))

        if adjuntar_archivos:
            for clave_ruta, nombre_adj in [
                ("word_path", "Documento_ULima.docx"),
                ("ppt_path", "Presentacion_ULima.pptx"),
                ("excel_path", "Hoja_Calculo_ULima.xlsx")
            ]:
                ruta = ent.get(clave_ruta)
                if ruta and os.path.exists(ruta):
                    with open(ruta, "rb") as f:
                        parte = MIMEBase("application", "octet-stream")
                        parte.set_payload(f.read())
                    encoders.encode_base64(parte)
                    parte.add_header("Content-Disposition", f"attachment; filename={nombre_adj}")
                    msg.attach(parte)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        gmail.users().messages().send(userId="me", body={"raw": raw}).execute()
        return f"Correo enviado exitosamente desde tu Gmail a {dest}."
    except Exception as e:
        return f"Error al enviar correo por Gmail API: {e}"


def crear_word_y_google_docs(titulo: str, contenido_markdown: str) -> str:
    """Crea un documento Word (.docx) descargable y, si Google está conectado, lo sube como Google Docs editable."""
    registrar_accion(f"📄 Creando Word y Google Docs: '{titulo}'")
    ruta_docx = f"/tmp/Doc_{int(time.time())}.docx"
    doc = Document()
    doc.add_heading(titulo, level=0)
    doc.add_paragraph(f"Generado por Agente IA ULima — {ahora_lima().strftime('%d/%m/%Y %H:%M')}")

    for linea in contenido_markdown.split("\n"):
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

    doc.save(ruta_docx)
    st.session_state["entregables_turno"]["word_path"] = ruta_docx
    st.session_state["entregables_turno"]["word_name"] = f"{titulo[:35].replace(' ', '_')}.docx"

    creds = st.session_state.get("google_creds")
    if creds:
        try:
            drive = build("drive", "v3", credentials=creds)
            meta = {"name": titulo, "mimeType": "application/vnd.google-apps.document"}
            media = MediaFileUpload(ruta_docx, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            creado = drive.files().create(body=meta, media_body=media, fields="webViewLink").execute()
            link = creado.get("webViewLink", "")
            st.session_state["entregables_turno"]["docs_link"] = link
            return f"Word creado y subido como Google Docs: {link}"
        except Exception as e:
            return f"Word (.docx) creado localmente (aviso Drive: {e})"
    return "Documento Word (.docx) creado y listo para descargar."


def crear_ppt_y_google_slides(titulo_portada: str, subtitulo: str, diapositivas_json: str) -> str:
    """Crea una presentación PowerPoint (.pptx) y Google Slides. `diapositivas_json` debe ser una lista JSON de objetos con keys 'titulo' y 'puntos' (lista de strings)."""
    registrar_accion(f"📊 Creando PowerPoint y Google Slides: '{titulo_portada}'")
    ruta_pptx = f"/tmp/PPT_{int(time.time())}.pptx"
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Portada oscura elegante
    portada = prs.slides.add_slide(prs.slide_layouts[0])
    bg = portada.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(20, 20, 24)

    if portada.shapes.title:
        portada.shapes.title.text = titulo_portada
        for p in portada.shapes.title.text_frame.paragraphs:
            p.font.color.rgb = RGBColor(249, 115, 22)
            p.font.bold = True
            p.font.size = Pt(36)
    if len(portada.placeholders) > 1:
        portada.placeholders[1].text = f"{subtitulo}\nAgente IA ULima — {ahora_lima().strftime('%d/%m/%Y')}"
        for p in portada.placeholders[1].text_frame.paragraphs:
            p.font.color.rgb = RGBColor(240, 240, 245)
            p.font.size = Pt(18)

    try:
        slides_data = json.loads(diapositivas_json)
        if not isinstance(slides_data, list):
            slides_data = [{"titulo": titulo_portada, "puntos": [str(diapositivas_json)]}]
    except Exception:
        slides_data = [{"titulo": titulo_portada, "puntos": [l.strip() for l in diapositivas_json.split("\n") if l.strip()][:6]}]

    for item in slides_data[:10]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        if slide.shapes.title:
            slide.shapes.title.text = str(item.get("titulo", "Contenido"))[:85]
            for p in slide.shapes.title.text_frame.paragraphs:
                p.font.color.rgb = RGBColor(234, 88, 12)
                p.font.bold = True
                p.font.size = Pt(28)
        if len(slide.placeholders) > 1:
            tf = slide.placeholders[1].text_frame
            tf.clear()
            puntos = item.get("puntos", [])
            if isinstance(puntos, str):
                puntos = [puntos]
            for idx, pt in enumerate(puntos[:6]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.text = f"• {str(pt)[:200]}"
                p.font.size = Pt(18)
                p.space_after = Pt(12)

    prs.save(ruta_pptx)
    st.session_state["entregables_turno"]["ppt_path"] = ruta_pptx
    st.session_state["entregables_turno"]["ppt_name"] = f"{titulo_portada[:35].replace(' ', '_')}.pptx"

    creds = st.session_state.get("google_creds")
    if creds:
        try:
            drive = build("drive", "v3", credentials=creds)
            meta = {"name": titulo_portada, "mimeType": "application/vnd.google-apps.presentation"}
            media = MediaFileUpload(ruta_pptx, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            creado = drive.files().create(body=meta, media_body=media, fields="webViewLink").execute()
            link = creado.get("webViewLink", "")
            st.session_state["entregables_turno"]["slides_link"] = link
            return f"PowerPoint creado y subido como Google Slides: {link}"
        except Exception as e:
            return f"PowerPoint (.pptx) creado localmente (aviso Drive: {e})"
    return "Presentación PowerPoint (.pptx) creada y lista para descargar."


def crear_excel_y_google_sheets(titulo_hoja: str, filas_json: str) -> str:
    """Crea un libro Excel (.xlsx) y Google Sheets. `filas_json` debe ser una lista JSON de listas (la primera sublista es el encabezado de columnas)."""
    registrar_accion(f"📈 Creando Excel y Google Sheets: '{titulo_hoja}'")
    ruta_xlsx = f"/tmp/Excel_{int(time.time())}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos ULima"

    header_fill = PatternFill(start_color="E85D04", end_color="E85D04", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)

    try:
        filas = json.loads(filas_json)
        if not isinstance(filas, list) or not filas:
            filas = [["Concepto", "Detalle"], [titulo_hoja, str(filas_json)]]
    except Exception:
        filas = [["Concepto", "Detalle"], [titulo_hoja, str(filas_json)]]

    for idx_r, fila in enumerate(filas):
        if not isinstance(fila, list):
            fila = [str(fila)]
        ws.append([str(c) for c in fila])
        if idx_r == 0:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            max_len = max(max_len, min(len(str(cell.value or "")), 55))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    wb.save(ruta_xlsx)
    st.session_state["entregables_turno"]["excel_path"] = ruta_xlsx
    st.session_state["entregables_turno"]["excel_name"] = f"{titulo_hoja[:35].replace(' ', '_')}.xlsx"

    creds = st.session_state.get("google_creds")
    if creds:
        try:
            drive = build("drive", "v3", credentials=creds)
            meta = {"name": titulo_hoja, "mimeType": "application/vnd.google-apps.spreadsheet"}
            media = MediaFileUpload(ruta_xlsx, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            creado = drive.files().create(body=meta, media_body=media, fields="webViewLink").execute()
            link = creado.get("webViewLink", "")
            st.session_state["entregables_turno"]["sheets_link"] = link
            return f"Excel creado y subido como Google Sheets: {link}"
        except Exception as e:
            return f"Excel (.xlsx) creado localmente (aviso Drive: {e})"
    return "Libro Excel (.xlsx) creado y listo para descargar."


HERRAMIENTAS_AGENTE = [
    buscar_en_internet,
    leer_pagina_web,
    ver_eventos_google_calendar,
    agendar_evento_google_calendar,
    leer_ultimos_correos_gmail,
    enviar_correo_gmail,
    crear_word_y_google_docs,
    crear_ppt_y_google_slides,
    crear_excel_y_google_sheets
]


# ==============================================================================
# PANTALLA 0: INICIO DE SESIÓN CON GOOGLE (TIPO GEMINI)
# ==============================================================================
if st.session_state["google_creds"] is None and not st.session_state["modo_sin_oauth"]:
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, centro, _ = st.columns([1, 1.3, 1])

    with centro:
        st.markdown("## 🎓 Agente IA — Universidad de Lima")
        st.markdown(
            "Conecta tu cuenta de Google para darle acceso a tu **Gmail, Google Calendar, Google Docs, Google Slides, Google Sheets** "
            "y **navegación por internet en tiempo real**."
        )

        if not st.session_state["api_key"]:
            api_tmp = st.text_input("🔑 Google AI Studio API Key", type="password", placeholder="AIzaSy...")
            if api_tmp.strip():
                st.session_state["api_key"] = api_tmp.strip()
                st.rerun()

        oauth_configurado = bool(st.session_state["client_id"] and st.session_state["client_secret"])

        if oauth_configurado and st.session_state["api_key"]:
            flujo_btn = crear_flujo_oauth()
            url_google, _ = flujo_btn.authorization_url(
                prompt="consent",
                access_type="offline",
                include_granted_scopes="true"
            )
            st.link_button(
                "🔵 Continuar con Google",
                url_google,
                type="primary",
                use_container_width=True
            )
        else:
            st.warning(
                "Para activar el botón oficial **Continuar con Google**, agrega `GEMINI_API_KEY`, `GOOGLE_CLIENT_ID` y `GOOGLE_CLIENT_SECRET` "
                "en **Settings → Secrets** de Streamlit."
            )

        st.markdown("")
        if st.button("⚡ Entrar ahora en Modo Rápido (Web + Word + PPT + Excel)", use_container_width=True):
            if not st.session_state["api_key"]:
                st.error("Ingresa primero tu Google AI Studio API Key arriba.")
            else:
                st.session_state["modo_sin_oauth"] = True
                st.rerun()

    st.stop()


# ==============================================================================
# INTERFAZ PRINCIPAL TIPO GEMINI (CHAT + ACCIONES AUTÓNOMAS EN VIVO)
# ==============================================================================
with st.sidebar:
    st.title("🎓 Agente IA ULima")
    if st.session_state["google_creds"]:
        st.success(f"🟢 {st.session_state['google_email']}")
        st.caption("Conectado a Gmail, Calendar y Drive")
    else:
        st.info("🟡 Modo Rápido (Sin OAuth)")
        st.caption("Web + Word + PPT + Excel activos")

    st.divider()
    st.markdown("**Herramientas del Agente:**")
    st.markdown("🌐 Navegador Web & Lectura URL")
    st.markdown("📅 Google Calendar (Leer y Agendar)")
    st.markdown("📧 Gmail (Leer y Enviar con adjuntos)")
    st.markdown("📄 Word + Google Docs")
    st.markdown("📊 PowerPoint + Google Slides")
    st.markdown("📈 Excel + Google Sheets")

    st.divider()
    if st.button("🗑️ Limpiar conversación", use_container_width=True):
        st.session_state["mensajes"] = []
        st.rerun()

    if st.button("🔄 Cambiar / Desconectar cuenta", use_container_width=True):
        st.session_state["google_creds"] = None
        st.session_state["google_email"] = ""
        st.session_state["modo_sin_oauth"] = False
        st.rerun()

st.markdown("### ✨ ¿Qué quieres que haga hoy por ti?")
st.caption("Pídeme lo que quieras en lenguaje natural: investigar en internet, orientarte entre las 14 carreras de la ULima, revisar/agendar fechas en Calendar, leer/enviar Gmails o crear Docs, PPT y Excel.")

# Botones rápidos de ejemplo (1 clic)
q1, q2, q3, q4 = st.columns(4)
prompt_rapido = None
if q1.button("🧭 Orientación 14 Carreras ULima", use_container_width=True):
    prompt_rapido = (
        "Investiga en internet sobre Ingeniería de Sistemas, Ingeniería Industrial y Administración en la Universidad de Lima. "
        "Compáralas en un Excel (Google Sheets) y hazme una presentación PPT (Google Slides) de 5 diapositivas para ayudarme a elegir."
    )
if q2.button("📅 Ver Calendar y Agendar Estudio", use_container_width=True):
    prompt_rapido = (
        "Revisa mi Google Calendar para los próximos días y agéndame una sesión de estudio de 90 minutos "
        "para mañana a las 6:00 pm titulada 'Repaso Examen ULima - ODS 4'."
    )
if q3.button("📊 Investigar Web + PPT + Word", use_container_width=True):
    prompt_rapido = (
        "Navega en internet sobre el impacto de la Inteligencia Artificial en la educación superior (ODS 4) en el Perú. "
        "Con esa información, créame un informe completo en Word (Google Docs) y una presentación PPT (Google Slides) de 6 diapositivas."
    )
if q4.button("📧 Revisar Gmail y Enviar Resumen", use_container_width=True):
    prompt_rapido = (
        "Revisa mis últimos 3 correos en Gmail, hazme un resumen ejecutivo de lo más importante y envíamelo "
        "a mi propio correo con el asunto 'Resumen Diario - Agente IA ULima'."
    )

# Mostrar historial del chat
for msg in st.session_state["mensajes"]:
    with st.chat_message(msg["role"]):
        if msg.get("acciones"):
            with st.expander("🛠️ Herramientas ejecutadas por el agente", expanded=False):
                for ac in msg["acciones"]:
                    st.code(ac, language="text")

        st.markdown(msg["content"])

        ent = msg.get("entregables", {})
        if ent:
            cols_links = []
            if ent.get("docs_link"):
                cols_links.append(("🟦 Abrir Google Docs", ent["docs_link"]))
            if ent.get("slides_link"):
                cols_links.append(("🟨 Abrir Google Slides", ent["slides_link"]))
            if ent.get("sheets_link"):
                cols_links.append(("🟩 Abrir Google Sheets", ent["sheets_link"]))
            if ent.get("calendar_link"):
                cols_links.append(("📅 Ver en Google Calendar", ent["calendar_link"]))
            if ent.get("gmail_link"):
                cols_links.append(("✉️ Abrir en Gmail", ent["gmail_link"]))

            if cols_links:
                c_ws = st.columns(len(cols_links))
                for idx_l, (label_l, url_l) in enumerate(cols_links):
                    c_ws[idx_l].link_button(label_l, url_l, use_container_width=True)

            cols_dl = []
            if ent.get("word_path") and os.path.exists(ent["word_path"]):
                cols_dl.append(("📄 Descargar Word (.docx)", ent["word_path"], ent.get("word_name", "Documento.docx"), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
            if ent.get("ppt_path") and os.path.exists(ent["ppt_path"]):
                cols_dl.append(("📊 Descargar PPT (.pptx)", ent["ppt_path"], ent.get("ppt_name", "Presentacion.pptx"), "application/vnd.openxmlformats-officedocument.presentationml.presentation"))
            if ent.get("excel_path") and os.path.exists(ent["excel_path"]):
                cols_dl.append(("📈 Descargar Excel (.xlsx)", ent["excel_path"], ent.get("excel_name", "Hoja.xlsx"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))

            if cols_dl:
                c_files = st.columns(len(cols_dl))
                for idx_f, (lbl_f, path_f, name_f, mime_f) in enumerate(cols_dl):
                    with open(path_f, "rb") as f_bin:
                        c_files[idx_f].download_button(lbl_f, data=f_bin.read(), file_name=name_f, mime=mime_f, use_container_width=True)

entrada_chat = st.chat_input("Escribe tu pedido aquí (ej. 'Investiga X, haz un PPT y Excel, agéndalo para el viernes a las 5 pm y mándalo por Gmail')...")
prompt_usuario = prompt_rapido or entrada_chat

if prompt_usuario:
    st.session_state["mensajes"].append({"role": "user", "content": prompt_usuario})
    with st.chat_message("user"):
        st.markdown(prompt_usuario)

    st.session_state["acciones_turno"] = []
    st.session_state["entregables_turno"] = {}

    ahora = ahora_lima()
    contexto_tiempo = f"{DIAS_ES[ahora.weekday()]} {ahora.strftime('%Y-%m-%d')} {ahora.strftime('%H:%M')} (Hora Lima, Perú)"
    correo_actual = st.session_state.get("google_email") or "No vinculado"

    instruccion_sistema = (
        f"Eres el Agente IA Autónomo de la Universidad de Lima (ULima - ODS 4), un asistente ejecutivo conectado a internet y a Google Workspace. "
        f"Fecha y hora actual exacta: {contexto_tiempo}. Cuenta Google conectada: {correo_actual}. "
        f"Carreras oficiales de la ULima (14): {', '.join(CARRERAS_ULIMA)}.\n\n"
        f"INSTRUCCIONES DE EJECUCIÓN:\n"
        f"1. Tienes herramientas reales para buscar en internet, leer URLs, leer/crear eventos en Google Calendar, leer/enviar correos en Gmail, "
        f"y crear archivos Word/Google Docs, PowerPoint/Google Slides y Excel/Google Sheets.\n"
        f"2. Cuando el usuario te pida hacer varias cosas (ej. investigar + crear PPT + crear Excel + agendar + enviar correo), "
        f"LLAMA A TODAS LAS HERRAMIENTAS NECESARIAS EN ORDEN antes de dar tu respuesta final (por ejemplo, crea primero los archivos Word/PPT/Excel y LUEGO llama a `enviar_correo_gmail` para que vayan adjuntos).\n"
        f"3. Para `crear_ppt_y_google_slides`, pasa `diapositivas_json` como JSON válido: "
        f'[{{"titulo": "Título Slide 1", "puntos": ["Idea 1", "Idea 2", "Idea 3"]}}, ...].\n'
        f"4. Para `crear_excel_y_google_sheets`, pasa `filas_json` como JSON válido de lista de listas: "
        f'[["Columna 1", "Columna 2"], ["Dato 1", "Dato 2"]].\n'
        f"5. Responde siempre en español claro, directo, bien estructurado y sin relleno innecesario."
    )

    with st.chat_message("assistant"):
        with st.status("⚡ Agente IA trabajando y ejecutando tus herramientas...", expanded=True) as status:
            try:
                client = genai.Client(api_key=st.session_state["api_key"])
                historial_texto = "\n".join(
                    [f"{'Usuario' if m['role'] == 'user' else 'Asistente'}: {m['content']}" for m in st.session_state["mensajes"][-6:]]
                )

                respuesta_texto = ""
                ultimo_err = ""
                for modelo_id in ["gemini-3.8-flash", "gemini-3.8-pro", "gemini-2.5-flash"]:
                    try:
                        resp = client.models.generate_content(
                            model=modelo_id,
                            contents=f"Historial reciente:\n{historial_texto}\n\nSolicitud actual del usuario: {prompt_usuario}",
                            config=types.GenerateContentConfig(
                                system_instruction=instruccion_sistema,
                                tools=HERRAMIENTAS_AGENTE,
                                temperature=0.3
                            )
                        )
                        respuesta_texto = resp.text or "Acciones completadas correctamente."
                        break
                    except Exception as e_mod:
                        ultimo_err = str(e_mod)
                        if "API_KEY_INVALID" in ultimo_err:
                            raise ValueError("La API Key de Gemini no es válida.")
                        time.sleep(1.5)

                if not respuesta_texto:
                    raise RuntimeError(f"Los servidores de Gemini están ocupados. Intenta de nuevo en unos segundos. ({ultimo_err[:100]})")

                status.update(label="✨ ¡Listo! Todas las tareas ejecutadas", state="complete", expanded=False)
            except Exception as err:
                status.update(label="⚠️ Ocurrió un aviso", state="error", expanded=True)
                respuesta_texto = f"No pude completar la acción: {err}"

        if st.session_state["acciones_turno"]:
            with st.expander("🛠️ Herramientas ejecutadas en vivo por el agente", expanded=True):
                for ac in st.session_state["acciones_turno"]:
                    st.code(ac, language="text")

        st.markdown(respuesta_texto)

        ent_actual = dict(st.session_state["entregables_turno"])
        if ent_actual:
            cols_links = []
            if ent_actual.get("docs_link"):
                cols_links.append(("🟦 Abrir Google Docs", ent_actual["docs_link"]))
            if ent_actual.get("slides_link"):
                cols_links.append(("🟨 Abrir Google Slides", ent_actual["slides_link"]))
            if ent_actual.get("sheets_link"):
                cols_links.append(("🟩 Abrir Google Sheets", ent_actual["sheets_link"]))
            if ent_actual.get("calendar_link"):
                cols_links.append(("📅 Ver en Google Calendar", ent_actual["calendar_link"]))
            if ent_actual.get("gmail_link"):
                cols_links.append(("✉️ Abrir en Gmail", ent_actual["gmail_link"]))

            if cols_links:
                c_ws = st.columns(len(cols_links))
                for idx_l, (label_l, url_l) in enumerate(cols_links):
                    c_ws[idx_l].link_button(label_l, url_l, use_container_width=True)

            cols_dl = []
            if ent_actual.get("word_path") and os.path.exists(ent_actual["word_path"]):
                cols_dl.append(("📄 Descargar Word (.docx)", ent_actual["word_path"], ent_actual.get("word_name", "Documento.docx"), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
            if ent_actual.get("ppt_path") and os.path.exists(ent_actual["ppt_path"]):
                cols_dl.append(("📊 Descargar PPT (.pptx)", ent_actual["ppt_path"], ent_actual.get("ppt_name", "Presentacion.pptx"), "application/vnd.openxmlformats-officedocument.presentationml.presentation"))
            if ent_actual.get("excel_path") and os.path.exists(ent_actual["excel_path"]):
                cols_dl.append(("📈 Descargar Excel (.xlsx)", ent_actual["excel_path"], ent_actual.get("excel_name", "Hoja.xlsx"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))

            if cols_dl:
                c_files = st.columns(len(cols_dl))
                for idx_f, (lbl_f, path_f, name_f, mime_f) in enumerate(cols_dl):
                    with open(path_f, "rb") as f_bin:
                        c_files[idx_f].download_button(lbl_f, data=f_bin.read(), file_name=name_f, mime=mime_f, use_container_width=True)

    st.session_state["mensajes"].append({
        "role": "assistant",
        "content": respuesta_texto,
        "acciones": list(st.session_state["acciones_turno"]),
        "entregables": ent_actual
    })
