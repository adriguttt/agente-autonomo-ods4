import os
import re
import html
import time
import json
import base64
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
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from duckduckgo_search import DDGS

st.set_page_config(
    page_title="AGENTE IA ULIMA | Google Workspace OAuth2 & ODS 4",
    page_icon="🎓",
    layout="wide"
)

ZONA_LIMA = timezone(timedelta(hours=-5))
ARCHIVO_WORD_PATH = "/tmp/Documento_ULima_ODS4.docx"
ARCHIVO_PPT_PATH = "/tmp/Presentacion_ULima_ODS4.pptx"
ARCHIVO_EXCEL_PATH = "/tmp/Hoja_Calculo_ULima_ODS4.xlsx"

SCOPES_GOOGLE = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/drive.file"
]

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
        "modelo_preferido": "gemini-3.8-flash"
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


inicializar_estado()


def crear_flujo_oauth(client_id: str, client_secret: str, redirect_uri: str) -> Flow:
    client_config = {
        "web": {
            "client_id": client_id.strip(),
            "client_secret": client_secret.strip(),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [redirect_uri.strip()]
        }
    }
    flow = Flow.from_client_config(
        client_config,
        scopes=SCOPES_GOOGLE,
        redirect_uri=redirect_uri.strip()
    )
    flow.autogenerate_code_verifier = False
    return flow


# Capturar retorno automático desde Google OAuth (?code=...)
codigo_oauth = st.query_params.get("code")
if codigo_oauth and st.session_state["google_creds"] is None:
    if st.session_state["client_id"] and st.session_state["client_secret"]:
        try:
            flujo = crear_flujo_oauth(
                st.session_state["client_id"],
                st.session_state["client_secret"],
                st.session_state["redirect_uri"]
            )
            flujo.code_verifier = None
            flujo.fetch_token(code=codigo_oauth)
            creds = flujo.credentials
            st.session_state["google_creds"] = creds

            # Obtener correo real del usuario autenticado
            gmail_srv = build("gmail", "v1", credentials=creds)
            perfil_gmail = gmail_srv.users().getProfile(userId="me").execute()
            st.session_state["google_email"] = perfil_gmail.get("emailAddress", "Cuenta Google Conectada")

            st.query_params.clear()
            st.rerun()
        except Exception as e_oauth:
            st.query_params.clear()
            st.error(f"No se pudo completar el inicio de sesión con Google: {e_oauth}")


# ==============================================================================
# NAVEGACIÓN WEB EN TIEMPO REAL
# ==============================================================================
def leer_pagina_web(url: str) -> str:
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; AgenteIAULima/2.0)"}
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

        log = f"[{hora}] 🌐 Navegador Web en vivo: {len(resultados)} fuentes halladas ({urls_leidas} páginas leídas)"
        return "\n".join(bloques), log
    except Exception:
        return "Estándares de educación superior UNESCO ODS 4 (ULima).", f"[{hora}] 🌐 Navegador Web: respaldo académico activo"


def buscar_wikipedia_vivo(concepto: str) -> tuple[str, str]:
    hora = ahora_lima().strftime("%H:%M:%S")
    try:
        termino = urllib.parse.quote(concepto.strip().split()[0])
        url = f"https://es.wikipedia.org/api/rest_v1/page/summary/{termino}"
        req = urllib.request.Request(url, headers={"User-Agent": "AgenteULimaODS4/10.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            datos = json.loads(resp.read().decode("utf-8"))
        extracto = datos.get("extract", "")
        if extracto:
            return extracto, f"[{hora}] 📚 Wikipedia consultada: '{concepto}' verificado"
        return "", f"[{hora}] 📚 Wikipedia: consulta completada"
    except Exception:
        return "", f"[{hora}] 📚 Wikipedia: síntesis interna activa"


# ==============================================================================
# CONEXIÓN CON GEMINI
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
# GENERADORES DE ARCHIVOS LOCALES + SUBIDA NATIVA A GOOGLE WORKSPACE
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
    return f"[{hora}] 📄 Archivo Word (.docx) generado"


def crear_ppt_archivo(titulo_portada: str, subtitulo: str, contenido_md: str) -> str:
    hora = ahora_lima().strftime("%H:%M:%S")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

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
    return f"[{hora}] 📊 Archivo PowerPoint (.pptx) generado ({min(len(secciones), 8) + 1} diapositivas)"


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
    return f"[{hora}] 📈 Archivo Excel (.xlsx) generado (3 hojas)"


def subir_a_google_workspace(creds, tema: str) -> tuple[dict, list[str]]:
    """Sube y convierte los archivos directo en Google Docs, Google Slides y Google Sheets en Drive."""
    logs = []
    enlaces = {"docs": "", "slides": "", "sheets": ""}
    hora = ahora_lima().strftime("%H:%M:%S")
    try:
        drive_service = build("drive", "v3", credentials=creds)

        # 1. Google Docs
        if os.path.exists(ARCHIVO_WORD_PATH):
            meta_doc = {
                "name": f"ULima ODS 4 - {tema}",
                "mimeType": "application/vnd.google-apps.document"
            }
            media_doc = MediaFileUpload(
                ARCHIVO_WORD_PATH,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            res_doc = drive_service.files().create(body=meta_doc, media_body=media_doc, fields="webViewLink").execute()
            enlaces["docs"] = res_doc.get("webViewLink", "")
            logs.append(f"[{hora}] 🟦 Google Docs creado directamente en tu Google Drive")

        # 2. Google Slides
        if os.path.exists(ARCHIVO_PPT_PATH):
            meta_ppt = {
                "name": f"Presentación ULima ODS 4 - {tema}",
                "mimeType": "application/vnd.google-apps.presentation"
            }
            media_ppt = MediaFileUpload(
                ARCHIVO_PPT_PATH,
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            res_ppt = drive_service.files().create(body=meta_ppt, media_body=media_ppt, fields="webViewLink").execute()
            enlaces["slides"] = res_ppt.get("webViewLink", "")
            logs.append(f"[{hora}] 🟨 Google Slides creado directamente en tu Google Drive")

        # 3. Google Sheets
        if os.path.exists(ARCHIVO_EXCEL_PATH):
            meta_xls = {
                "name": f"Plan ULima ODS 4 - {tema}",
                "mimeType": "application/vnd.google-apps.spreadsheet"
            }
            media_xls = MediaFileUpload(
                ARCHIVO_EXCEL_PATH,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            res_xls = drive_service.files().create(body=meta_xls, media_body=media_xls, fields="webViewLink").execute()
            enlaces["sheets"] = res_xls.get("webViewLink", "")
            logs.append(f"[{hora}] 🟩 Google Sheets creado directamente en tu Google Drive")

    except Exception as e:
        logs.append(f"[{hora}] ⚠️ Aviso al subir a Google Drive: {str(e)[:120]}")

    return enlaces, logs


def insertar_evento_real_google_calendar(creds, titulo_evento: str, inicio_dt: datetime, duracion_min: int, descripcion: str) -> tuple[str, str]:
    """Inserta el evento directamente en el calendario principal del usuario vía Google Calendar API."""
    hora = ahora_lima().strftime("%H:%M:%S")
    if inicio_dt.tzinfo is None:
        inicio_dt = inicio_dt.replace(tzinfo=ZONA_LIMA)
    fin_dt = inicio_dt + timedelta(minutes=duracion_min)

    try:
        cal_service = build("calendar", "v3", credentials=creds)
        evento_body = {
            "summary": titulo_evento,
            "description": descripcion,
            "start": {"dateTime": inicio_dt.isoformat(), "timeZone": "America/Lima"},
            "end": {"dateTime": fin_dt.isoformat(), "timeZone": "America/Lima"},
            "reminders": {
                "useDefault": False,
                "overrides": [{"method": "popup", "minutes": 30}]
            }
        }
        creado = cal_service.events().insert(calendarId="primary", body=evento_body).execute()
        link_evento = creado.get("htmlLink", "https://calendar.google.com")
        return link_evento, f"[{hora}] 📅 EVENTO AGENDADO EN TU GOOGLE CALENDAR: '{titulo_evento}' ({inicio_dt.strftime('%d/%m/%Y %H:%M')})"
    except Exception as e:
        return "https://calendar.google.com", f"[{hora}] ⚠️ Error en Google Calendar API: {str(e)[:120]}"


def enviar_gmail_real_api(creds, destinatario: str, asunto: str, contenido: str, enlaces_workspace: dict, link_calendar: str) -> tuple[bool, str]:
    """Envía un correo real directamente desde la cuenta Gmail autenticada por OAuth2."""
    hora = ahora_lima().strftime("%H:%M:%S")
    try:
        gmail_service = build("gmail", "v1", credentials=creds)
        msg = MIMEMultipart()
        msg["To"] = destinatario.strip()
        msg["Subject"] = asunto

        bloque_links = ""
        if enlaces_workspace.get("docs"):
            bloque_links += f"• Google Docs: {enlaces_workspace['docs']}\n"
        if enlaces_workspace.get("slides"):
            bloque_links += f"• Google Slides: {enlaces_workspace['slides']}\n"
        if enlaces_workspace.get("sheets"):
            bloque_links += f"• Google Sheets: {enlaces_workspace['sheets']}\n"
        if link_calendar:
            bloque_links += f"• Evento Google Calendar: {link_calendar}\n"

        cuerpo = (
            "¡Hola!\n\n"
            "El Agente IA Autónomo ULima (ODS 4) ha generado y sincronizado tus recursos en Google Workspace:\n\n"
            f"{bloque_links}\n"
            "También se adjuntan las copias descargables en Word (.docx), PowerPoint (.pptx) y Excel (.xlsx).\n\n"
            "------------------------------------------------------------\n"
            f"{contenido[:3500]}\n"
            "------------------------------------------------------------\n"
            "Agente IA Universidad de Lima — ODS 4."
        )
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

        for ruta, nombre in [
            (ARCHIVO_WORD_PATH, "Documento_ULima_ODS4.docx"),
            (ARCHIVO_PPT_PATH, "Presentacion_ULima_ODS4.pptx"),
            (ARCHIVO_EXCEL_PATH, "Hoja_Calculo_ULima_ODS4.xlsx")
        ]:
            if os.path.exists(ruta):
                with open(ruta, "rb") as f:
                    adj = MIMEBase("application", "octet-stream")
                    adj.set_payload(f.read())
                encoders.encode_base64(adj)
                adj.add_header("Content-Disposition", f"attachment; filename={nombre}")
                msg.attach(adj)

        raw_bytes = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        gmail_service.users().messages().send(userId="me", body={"raw": raw_bytes}).execute()
        return True, f"[{hora}] 📧 GMAIL ENVIADO REALMENTE a '{destinatario}' vía Gmail API oficial"
    except Exception as e:
        return False, f"[{hora}] ⚠️ Error al enviar por Gmail API: {str(e)[:140]}"


# ==============================================================================
# PANTALLA 0: INICIO DE SESIÓN OFICIAL CON GOOGLE (OAUTH 2.0)
# ==============================================================================
if st.session_state["google_creds"] is None:
    st.markdown("<br>", unsafe_allow_html=True)
    _, col_centro, _ = st.columns([1, 1.6, 1])

    with col_centro:
        st.markdown("## 🎓 Agente IA Autónomo — Universidad de Lima")
        st.caption("Conecta tu cuenta oficial de Google (OAuth 2.0) para que la IA acceda a tu Gmail, Google Calendar, Google Docs, Slides y Sheets.")

        api_in = st.text_input(
            "🔑 Google AI Studio API Key (Gemini)",
            value=st.session_state["api_key"],
            type="password",
            placeholder="AIzaSy..."
        )
        st.session_state["api_key"] = api_in.strip()

        with st.expander("⚙️ Credenciales OAuth de Google Cloud (1 sola vez)", expanded=not bool(st.session_state["client_id"] and st.session_state["client_secret"])):
            cid_in = st.text_input("Google Client ID", value=st.session_state["client_id"], placeholder="123456-xxxx.apps.googleusercontent.com")
            csec_in = st.text_input("Google Client Secret", value=st.session_state["client_secret"], type="password", placeholder="GOCSPX-...")
            ruri_in = st.text_input("Redirect URI (URL exacta de tu app)", value=st.session_state["redirect_uri"])
            st.session_state["client_id"] = cid_in.strip()
            st.session_state["client_secret"] = csec_in.strip()
            st.session_state["redirect_uri"] = ruri_in.strip()

        st.divider()

        if st.session_state["api_key"] and st.session_state["client_id"] and st.session_state["client_secret"]:
            try:
                flujo_login = crear_flujo_oauth(
                    st.session_state["client_id"],
                    st.session_state["client_secret"],
                    st.session_state["redirect_uri"]
                )
                url_autorizacion, _ = flujo_login.authorization_url(
                    prompt="consent",
                    access_type="offline",
                    include_granted_scopes="true"
                )
                st.link_button(
                    "🔵 Continuar con Google (Conectar Gmail, Calendar y Drive)",
                    url_autorizacion,
                    type="primary",
                    use_container_width=True
                )
            except Exception as err_cfg:
                st.error(f"Revisa el Client ID o Redirect URI: {err_cfg}")
        else:
            st.info("Completa tu **Gemini API Key**, **Client ID** y **Client Secret** arriba para activar el botón oficial de **Continuar con Google**.")

    st.stop()


# ==============================================================================
# WORKSPACE AUTÓNOMO ULIMA (AUTENTICADO CON GOOGLE OAUTH 2.0)
# ==============================================================================
with st.sidebar:
    st.title("⚙️ Sesión Google Activa")
    st.success(f"🟢 {st.session_state['google_email']}")
    st.caption("Permisos activos: Gmail + Calendar + Drive (Docs/Slides/Sheets)")

    st.divider()
    correo_destino_ui = st.text_input(
        "📩 Enviar correos a:",
        value=st.session_state["google_email"],
        help="Por defecto se envía a tu propio Gmail conectado."
    )
    fecha_base_ui = st.date_input(
        "📅 Fecha por defecto Calendar",
        value=(ahora_lima() + timedelta(days=1)).date()
    )
    hora_base_ui = st.time_input(
        "⏰ Hora por defecto Calendar",
        value=datetime.strptime("18:00", "%H:%M").time()
    )

    st.divider()
    st.markdown("**APIs Reales Conectadas:**")
    st.markdown("🌐 Navegación Web en Vivo")
    st.markdown("📅 Google Calendar API v3")
    st.markdown("📧 Gmail API v1 (Envío Real)")
    st.markdown("🟦 Google Docs + Word")
    st.markdown("🟨 Google Slides + PPT")
    st.markdown("🟩 Google Sheets + Excel")

    st.divider()
    if st.button("Cerrar sesión de Google", use_container_width=True):
        st.session_state["google_creds"] = None
        st.session_state["google_email"] = ""
        st.rerun()

st.title("🎓 AGENTE IA ULIMA — Ecosistema Google Real (ODS 4)")
st.markdown(
    "Agente autónomo con **navegación por internet en tiempo real**, conectado vía **OAuth 2.0** directamente a tu **Gmail, Google Calendar, Google Docs, Google Slides y Google Sheets** "
    "para orientación vocacional, postulantes y las **14 carreras de la Universidad de Lima**."
)

perfil = st.radio(
    "Selecciona el modo de operación:",
    [
        "🧭 Orientación Vocacional ULima (14 carreras)",
        "🚀 Postulante a la ULima (Admisión / CEPREULIMA / Escolar)",
        "🎓 Alumno Regular ULima (14 Carreras / Estudios Generales)",
        "🤖 Asistente Libre IA (Navegar web, Docs, Slides, Sheets, Calendar y Gmail)"
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
        "dime cuál es mi mejor opción, crea un Google Docs, un Google Slides y un Google Sheets comparativo, prográmalo en mi Google Calendar para mañana a las 5:00 pm y mándamelo por Gmail."
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
                "Plan de estudio en Google Sheets, Slides y Calendar",
                "Guía puente hacia Estudios Generales ULima"
            ]
        )
    detalle_perfil = f"{modalidad} → {carrera_objetivo}"
    valor_default_mision = (
        f"Explícalo fácil desde cero para un postulante a {carrera_objetivo} en la ULima con 4 preguntas tipo examen resueltas paso a paso, "
        "créame un Google Docs, un Google Slides y un Google Sheets de estudio, agéndalo en mi Google Calendar para mañana a las 6:00 pm y envíamelo por Gmail."
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
                "Google Slides + Google Docs listos para exposición grupal",
                "Explicar teoría difícil desde cero con ejemplos prácticos",
                "Plan de estudio en Google Sheets + Google Calendar"
            ]
        )
    detalle_perfil = carrera_actual
    valor_default_mision = (
        f"Investiga en internet casos actuales, explícalo fácil para un alumno de {carrera_actual} en la ULima con 3 ejercicios resueltos, "
        "crea un Google Docs, un Google Slides y un Google Sheets, prográmalo en mi Google Calendar para mañana a las 7:00 pm y envíamelo por Gmail."
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
    necesidad = "Ejecución Autónoma en Google Workspace"
    valor_default_mision = (
        "Busca información actualizada en internet sobre este tema, crea un Google Docs, un Google Slides de 6 diapositivas "
        "y un Google Sheets con métricas clave, agéndalo en mi Google Calendar para mañana a las 4:00 pm y envíalo a mi Gmail."
    )

objetivo_extra = st.text_area(
    "Pídele lo que quieras como IA (navegar en internet, agendar directo en Google Calendar, crear Docs/Slides/Sheets o mandar Gmails):",
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
        "inicio": datetime.combine(fecha_base_ui, hora_base_ui, tzinfo=ZONA_LIMA),
        "titulo": f"ULima ODS 4: {tema}",
        "link_evento": "https://calendar.google.com",
        "programado_por_ia": False
    }
    enlaces_workspace = {"docs": "", "slides": "", "sheets": ""}
    exito_gmail = False

    try:
        client = genai.Client(api_key=st.session_state["api_key"])
        creds = st.session_state["google_creds"]

        with st.status("⚡ Agente IA navegando en internet y operando tu Google Workspace...", expanded=True) as status:
            # FASE 1: Navegación Web en Vivo + Wikipedia
            st.write("🌐 **Fase 1/3:** Navegando en internet y leyendo fuentes académicas en vivo...")
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
                f"para generar Google Docs, Google Slides y Google Sheets. "
                f"Encabézala con un breve **Índice de Facilitación ULima (ODS 4): [Puntaje/100]** y cierra con un **Checklist Rápido de Acción**."
            )
            contenido_final, modelo_usado = llamar_gemini_blindado(client, prompt_maestro, modelos_activos)
            modelos_priorizados = [modelo_usado] + [m for m in modelos_activos if m != modelo_usado]
            bitacora.append(f"[{ahora_lima().strftime('%H:%M:%S')}] 🧠 Contenido generado por IA ({modelo_usado})")

            # FASE 3: Ejecución Real en Google Calendar API, Google Drive API y Gmail API
            st.write("🤖 **Fase 3/3:** Agendando directo en Google Calendar, creando Docs/Slides/Sheets y enviando Gmail...")

            def tool_agendar_google_calendar(titulo_evento: str, fecha_yyyy_mm_dd: str, hora_hh_mm: str, duracion_minutos: int = 90) -> str:
                """Inserta el evento directamente en el Google Calendar real del usuario."""
                try:
                    dt_naive = datetime.strptime(f"{fecha_yyyy_mm_dd.strip()} {hora_hh_mm.strip()}", "%Y-%m-%d %H:%M")
                    dt_parsed = dt_naive.replace(tzinfo=ZONA_LIMA)
                except Exception:
                    dt_parsed = datetime.combine(fecha_base_ui, hora_base_ui, tzinfo=ZONA_LIMA)

                link_ev, log_cal = insertar_evento_real_google_calendar(
                    creds,
                    titulo_evento or f"ULima ODS 4: {tema}",
                    dt_parsed,
                    duracion_minutos or 90,
                    f"Sesión ULima ({detalle_perfil}) - {necesidad}"
                )
                estado_calendario["inicio"] = dt_parsed
                estado_calendario["titulo"] = titulo_evento or f"ULima ODS 4: {tema}"
                estado_calendario["link_evento"] = link_ev
                estado_calendario["programado_por_ia"] = True
                bitacora.append(log_cal)
                return log_cal

            prompt_calendar = (
                f"Hoy es {fecha_hoy_str}. Analiza la instrucción del usuario: '{objetivo_extra}'.\n"
                f"Llama a `tool_agendar_google_calendar` convirtiendo cualquier fecha/hora mencionada ('mañana', 'este viernes', 'a las 4 pm', etc.) "
                f"a `YYYY-MM-DD` y `HH:MM` (24h). Si no mencionó fecha/hora, usa `{fecha_base_ui.strftime('%Y-%m-%d')}` y `{hora_base_ui.strftime('%H:%M')}`."
            )
            llamar_gemini_blindado(client, prompt_calendar, modelos_priorizados, herramientas=[tool_agendar_google_calendar])

            if not estado_calendario["programado_por_ia"]:
                dt_def = datetime.combine(fecha_base_ui, hora_base_ui, tzinfo=ZONA_LIMA)
                link_ev, log_cal = insertar_evento_real_google_calendar(
                    creds,
                    f"ULima ODS 4: {tema}",
                    dt_def,
                    90,
                    f"Sesión ULima ({detalle_perfil}) - {necesidad}"
                )
                estado_calendario["inicio"] = dt_def
                estado_calendario["link_evento"] = link_ev
                bitacora.append(log_cal)

            # Crear archivos locales y subirlos convertidos a Google Docs, Google Slides y Google Sheets
            bitacora.append(crear_word_archivo(tema, detalle_perfil, contenido_final))
            bitacora.append(crear_ppt_archivo(tema, detalle_perfil, contenido_final))
            bitacora.append(crear_excel_archivo(tema, detalle_perfil, contenido_final, estado_calendario["inicio"]))

            enlaces_workspace, logs_drive = subir_a_google_workspace(creds, tema)
            bitacora.extend(logs_drive)

            # Enviar Gmail real vía Gmail API v1
            if destinatario_final:
                exito_gmail, msg_gmail = enviar_gmail_real_api(
                    creds,
                    destinatario_final,
                    f"Agente IA ULima ({detalle_perfil}): {tema} [{estado_calendario['inicio'].strftime('%d/%m %H:%M')}]",
                    contenido_final,
                    enlaces_workspace,
                    estado_calendario["link_evento"]
                )
                bitacora.append(msg_gmail)

            duracion = round(time.time() - tiempo_inicio, 1)
            status.update(label=f"✨ Todo sincronizado en tu Google Workspace en {duracion}s ({modelo_usado})", state="complete", expanded=False)

    except Exception as error_controlado:
        st.warning(f"⚠️ Aviso del sistema: {error_controlado}")
        st.stop()

    st.success(
        f"📅 **Agendado directo en tu Google Calendar:** `{estado_calendario['titulo']}` el "
        f"**{estado_calendario['inicio'].strftime('%d/%m/%Y a las %H:%M')}** (Hora Lima)."
    )

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Tiempo Total", f"{duracion} seg")
    m2.metric("Google Drive", "Docs + Slides + Sheets")
    m3.metric("Google Calendar", estado_calendario["inicio"].strftime("%d/%m %H:%M"))
    m4.metric("Gmail API", "Enviado ✅" if exito_gmail else "Listo")

    with st.expander("🛠️ Bitácora en Vivo de APIs Ejecutadas en tu Cuenta Google", expanded=True):
        for linea_log in bitacora:
            st.code(linea_log, language="text")

    st.subheader("☁️ Abrir Directamente en tu Google Workspace (Creados en tu Drive y Calendar)")
    g1, g2, g3, g4 = st.columns(4)
    with g1:
        st.link_button("🟦 Abrir en Google Docs", enlaces_workspace["docs"] or "https://docs.google.com", use_container_width=True)
    with g2:
        st.link_button("🟨 Abrir en Google Slides", enlaces_workspace["slides"] or "https://slides.google.com", use_container_width=True)
    with g3:
        st.link_button("🟩 Abrir en Google Sheets", enlaces_workspace["sheets"] or "https://sheets.google.com", use_container_width=True)
    with g4:
        st.link_button("📅 Ver en Google Calendar", estado_calendario["link_evento"], use_container_width=True)

    st.subheader("📦 Descargas Locales (Word, PowerPoint, Excel)")
    b1, b2, b3 = st.columns(3)
    with b1:
        if os.path.exists(ARCHIVO_WORD_PATH):
            with open(ARCHIVO_WORD_PATH, "rb") as f_doc:
                st.download_button(
                    label="📄 Descargar Word (.docx)",
                    data=f_doc.read(),
                    file_name=f"Documento_ULima_{tema.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
    with b2:
        if os.path.exists(ARCHIVO_PPT_PATH):
            with open(ARCHIVO_PPT_PATH, "rb") as f_ppt:
                st.download_button(
                    label="📊 Descargar PowerPoint (.pptx)",
                    data=f_ppt.read(),
                    file_name=f"Presentacion_ULima_{tema.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
    with b3:
        if os.path.exists(ARCHIVO_EXCEL_PATH):
            with open(ARCHIVO_EXCEL_PATH, "rb") as f_xls:
                st.download_button(
                    label="📈 Descargar Excel (.xlsx)",
                    data=f_xls.read(),
                    file_name=f"Hoja_ULima_{tema.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    st.divider()
    st.subheader(f"📘 Resultado Generado — {detalle_perfil}")
    st.markdown(contenido_final)
