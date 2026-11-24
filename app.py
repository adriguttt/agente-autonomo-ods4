import os
import re
import time
import socket

os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"] = "1"
socket.setdefaulttimeout(15)

import io
import base64
import datetime
from zoneinfo import ZoneInfo
from urllib.parse import quote_plus
from email.message import EmailMessage

import requests
import streamlit as st
from bs4 import BeautifulSoup
from duckduckgo_search import DDGS
from docx import Document
from pptx import Presentation
from openpyxl import Workbook

from google import genai
from google.genai import types
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

# ==============================================================================
# 1. CONFIGURACIÓN GENERAL Y MULTI-TENANT (ODS 4 - UNIVERSIDAD DE LIMA)
# ==============================================================================
st.set_page_config(
    page_title="Agente IA — Universidad de Lima (ODS 4)",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/tasks",
    "https://www.googleapis.com/auth/contacts.readonly",
    "https://www.googleapis.com/auth/forms.body",
]

REQUIRED_WORKSPACE_SCOPES = {
    "https://www.googleapis.com/auth/gmail.modify": "Gmail",
    "https://www.googleapis.com/auth/calendar": "Calendar",
    "https://www.googleapis.com/auth/drive.file": "Drive (Docs/Slides/Sheets)",
    "https://www.googleapis.com/auth/tasks": "Tasks",
    "https://www.googleapis.com/auth/contacts.readonly": "Contactos",
    "https://www.googleapis.com/auth/forms.body": "Forms",
}

MODELOS_DISPONIBLES = [
    "gemini-3.8-flash",
    "gemini-3.5-flash-lite",
]


def get_gemini_keys() -> list[str]:
    claves = []
    manual = st.session_state.get("custom_gemini_key", "").strip()
    if manual:
        claves.append(manual)
    secret_raw = str(st.secrets.get("GEMINI_API_KEY", "")).strip()
    for k in secret_raw.split(","):
        k_clean = k.strip()
        if k_clean and k_clean not in claves:
            claves.append(k_clean)
    return claves


def log_tool(nombre: str) -> None:
    st.session_state.setdefault("tool_logs", []).append(nombre)


def get_system_instruction() -> str:
    ahora_lima = datetime.datetime.now(ZoneInfo("America/Lima")).strftime(
        "%Y-%m-%d %H:%M (%A)"
    )
    usuario = st.session_state.get("user_email", "Modo Rápido (sin cuenta Google)")
    return f"""
Eres el Agente IA Autónomo de la Universidad de Lima enfocado en el ODS 4 (Educación de Calidad).
Ayudas a estudiantes, docentes y postulantes a investigar, estudiar, gestionar su tiempo y automatizar tareas reales.
Usuario conectado actualmente: {usuario}.

Conocimiento base ULima (14 carreras de pregrado):
1. Administración
2. Arquitectura
3. Comunicación
4. Contabilidad y Finanzas
5. Derecho
6. Economía
7. Ingeniería Ambiental
8. Ingeniería Civil
9. Ingeniería de Sistemas
10. Ingeniería Industrial
11. Ingeniería Mecatrónica
12. Marketing
13. Negocios Internacionales
14. Psicología

Reglas estrictas de ejecución:
- Ejecuta las herramientas disponibles de forma autónoma y encadenada cuando el usuario pida varias acciones en un mismo mensaje.
- JAMÁS afirmes que enviaste un correo, creaste un archivo, agendaste un evento o creaste una tarea/encuesta si la herramienta no devolvió una confirmación explícita ('ENVIADO_OK', 'ARCHIVO_OK', 'EVENTO_OK', 'TAREA_OK', 'FORM_OK') o si devolvió 'ERROR:'. Si hay un error, explica el motivo exacto al usuario.
- Para enviar correos usa gestionar_gmail con accion='enviar' y un correo electrónico válido en destinatario. Si solo te dan un nombre (ej. "Carlos"), consulta primero buscar_contactos_google.
- Si el usuario pide crear un documento, presentación o Excel y enviarlo por Gmail, genera primero el archivo y luego llama a gestionar_gmail con adjuntar_ultimo_archivo=True.
- Responde siempre en español claro, estructurado y con enlaces directos a los recursos creados.
- Fecha y hora actual en Lima, Perú: {ahora_lima}
"""


def get_oauth_flow() -> Flow:
    client_config = {
        "web": {
            "client_id": st.secrets["GOOGLE_CLIENT_ID"],
            "client_secret": st.secrets["GOOGLE_CLIENT_SECRET"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [st.secrets["REDIRECT_URI"]],
        }
    }
    try:
        return Flow.from_client_config(
            client_config=client_config,
            scopes=SCOPES,
            redirect_uri=st.secrets["REDIRECT_URI"],
            autogenerate_code_verifier=False,
        )
    except TypeError:
        return Flow.from_client_config(
            client_config=client_config,
            scopes=SCOPES,
            redirect_uri=st.secrets["REDIRECT_URI"],
        )


def get_user_credentials() -> Credentials | None:
    creds_dict = st.session_state.get("google_creds")
    if not creds_dict:
        return None
    return Credentials.from_authorized_user_info(creds_dict)


def check_scope(creds: Credentials, required_scope: str) -> bool:
    granted = set(creds.scopes or [])
    if not granted:
        return True
    return required_scope in granted


def format_google_error(servicio: str, err: Exception) -> str:
    if isinstance(err, HttpError):
        status = getattr(err.resp, "status", "?")
        if status == 403:
            return (
                f"ERROR: Permiso denegado en {servicio} (403). Verifica que marcaste todas las casillas "
                f"al iniciar sesión con Google y que la API de {servicio} esté habilitada en Google Cloud."
            )
        if status == 401:
            return "ERROR: La sesión de Google expiró (401). Cierra sesión y vuelve a entrar."
    return f"ERROR en {servicio}: {err}"


# ==============================================================================
# 2. HERRAMIENTAS AUTÓNOMAS DEL AGENTE (AISLADAS POR SESIÓN)
# ==============================================================================
def buscar_internet(consulta: str) -> str:
    """Busca información actualizada en internet y devuelve títulos, enlaces y resúmenes."""
    log_tool(f"🌐 Web: {consulta}")
    try:
        resultados = []
        with DDGS(timeout=6) as ddgs:
            for r in ddgs.text(consulta, max_results=5):
                resultados.append(
                    f"- {r.get('title')}: {r.get('body')} (URL: {r.get('href')})"
                )
        if resultados:
            return "\n".join(resultados)
    except Exception:
        pass

    try:
        resp = requests.get(
            "https://html.duckduckgo.com/html/",
            params={"q": consulta},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=6,
        )
        soup = BeautifulSoup(resp.text, "html.parser")
        items = []
        for res in soup.select(".result")[:5]:
            titulo = res.select_one(".result__a")
            snippet = res.select_one(".result__snippet")
            if titulo:
                items.append(
                    f"- {titulo.get_text(strip=True)}: {snippet.get_text(strip=True) if snippet else ''} (URL: {titulo.get('href', '')})"
                )
        if items:
            return "\n".join(items)
    except Exception:
        pass

    return "Sin resultados externos en este momento; responde usando tu conocimiento académico actualizado."


def leer_pagina_web(url: str) -> str:
    """Extrae el texto principal de una URL pública para analizar artículos o documentos web."""
    log_tool(f"🔗 URL: {url}")
    try:
        resp = requests.get(url, timeout=8, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer"]):
            tag.decompose()
        texto = " ".join(soup.stripped_strings)
        return texto[:6000] if texto else "La página no contiene texto legible."
    except Exception as e:
        return f"ERROR: No se pudo leer la URL: {e}"


def buscar_videos_youtube(tema: str, max_resultados: int = 5) -> str:
    """Busca videos educativos, tutoriales o conferencias en YouTube sobre un tema académico."""
    log_tool(f"▶️ YouTube: {tema}")
    limite = min(max(1, max_resultados), 8)
    claves = get_gemini_keys()
    if claves:
        try:
            yt = build("youtube", "v3", developerKey=claves[0])
            res = (
                yt.search()
                .list(
                    part="snippet",
                    q=tema,
                    type="video",
                    maxResults=limite,
                    relevanceLanguage="es",
                )
                .execute()
            )
            items = res.get("items", [])
            if items:
                return "\n".join(
                    [
                        f"- **{it['snippet']['title']}** (Canal: {it['snippet']['channelTitle']}) -> https://www.youtube.com/watch?v={it['id']['videoId']}"
                        for it in items
                    ]
                )
        except Exception:
            pass

    try:
        resultados = []
        with DDGS(timeout=6) as ddgs:
            for r in ddgs.text(f"site:youtube.com/watch {tema}", max_results=limite):
                resultados.append(f"- **{r.get('title')}** -> {r.get('href')}")
        if resultados:
            return "\n".join(resultados)
    except Exception:
        pass

    query_yt = quote_plus(tema)
    return f"Búsqueda directa de videos educativos en YouTube sobre **{tema}**: https://www.youtube.com/results?search_query={query_yt}"


def gestionar_google_tasks(
    accion: str, titulo: str = "", notas: str = "", fecha_limite: str = ""
) -> str:
    """
    Lista o crea tareas en Google Tasks.
    accion: 'listar' o 'crear'.
    fecha_limite: opcional en formato 'YYYY-MM-DD'.
    """
    log_tool(f"✅ Tasks ({accion})")
    creds = get_user_credentials()
    if not creds:
        return "ERROR: Modo Rápido activo. Conecta tu cuenta de Google para usar Google Tasks."
    if not check_scope(creds, "https://www.googleapis.com/auth/tasks"):
        return "ERROR: Falta el permiso de Google Tasks. Cierra sesión y vuelve a entrar marcando todas las casillas."
    try:
        service = build("tasks", "v1", credentials=creds)
        if accion.lower().strip() in ["crear", "create", "agregar", "add"] and titulo:
            body = {"title": titulo.strip(), "notes": notas}
            if fecha_limite and re.match(r"^\d{4}-\d{2}-\d{2}$", fecha_limite.strip()):
                body["due"] = f"{fecha_limite.strip()}T23:59:00.000Z"
            creada = service.tasks().insert(tasklist="@default", body=body).execute()
            return f"TAREA_OK: Tarea creada en Google Tasks: '{creada.get('title')}' (ID: {creada.get('id')})."
        tareas = (
            service.tasks()
            .list(tasklist="@default", showCompleted=False, maxResults=15)
            .execute()
            .get("items", [])
        )
        if not tareas:
            return "No tienes tareas pendientes en Google Tasks."
        return "\n".join(
            [
                f"- {t.get('title')} (Vence: {t.get('due', 'Sin fecha')[:10]})"
                for t in tareas
            ]
        )
    except Exception as e:
        return format_google_error("Google Tasks", e)


def buscar_contactos_google(nombre_o_correo: str) -> str:
    """Busca personas en Google Contactos del usuario para obtener sus correos o teléfonos."""
    log_tool(f"👥 Contactos: {nombre_o_correo}")
    creds = get_user_credentials()
    if not creds:
        return "ERROR: Modo Rápido activo. Conecta tu cuenta de Google para buscar en Contactos."
    if not check_scope(creds, "https://www.googleapis.com/auth/contacts.readonly"):
        return "ERROR: Falta el permiso de Google Contactos. Cierra sesión y vuelve a entrar marcando todas las casillas."
    try:
        people = build("people", "v1", credentials=creds)
        res = (
            people.people()
            .searchContacts(
                query=nombre_o_correo, readMask="names,emailAddresses,phoneNumbers"
            )
            .execute()
        )
        results = res.get("results", [])
        if not results:
            return f"No se encontraron contactos que coincidan con '{nombre_o_correo}'."
        salida = []
        for r in results:
            p = r.get("person", {})
            nombre = (p.get("names") or [{}])[0].get("displayName", "Sin nombre")
            correo = (p.get("emailAddresses") or [{}])[0].get("value", "Sin correo")
            tel = (p.get("phoneNumbers") or [{}])[0].get("value", "Sin teléfono")
            salida.append(f"- {nombre} | Correo: {correo} | Tel: {tel}")
        return "\n".join(salida)
    except Exception as e:
        return format_google_error("Google Contactos", e)


def crear_formulario_google(
    titulo: str, descripcion: str, preguntas_separadas_por_punto_y_coma: str
) -> str:
    """
    Crea un formulario o encuesta real en Google Forms con preguntas de respuesta corta.
    preguntas_separadas_por_punto_y_coma: ejemplo '¿Cuál es tu carrera?;¿Qué tema prefieres?'
    """
    log_tool(f"📝 Forms: {titulo}")
    creds = get_user_credentials()
    if not creds:
        return "ERROR: Modo Rápido activo. Conecta tu cuenta de Google para crear Google Forms."
    if not check_scope(creds, "https://www.googleapis.com/auth/forms.body"):
        return "ERROR: Falta el permiso de Google Forms. Cierra sesión y vuelve a entrar marcando todas las casillas."
    try:
        forms = build("forms", "v1", credentials=creds)
        form_creado = (
            forms.forms()
            .create(body={"info": {"title": titulo, "documentTitle": titulo}})
            .execute()
        )
        form_id = form_creado["formId"]
        requests_batch = []
        if descripcion:
            requests_batch.append(
                {
                    "updateFormInfo": {
                        "info": {"description": descripcion},
                        "updateMask": "description",
                    }
                }
            )
        raw_pregs = preguntas_separadas_por_punto_y_coma
        if ";" in raw_pregs:
            preguntas = [p.strip() for p in raw_pregs.split(";") if p.strip()]
        elif "|" in raw_pregs:
            preguntas = [p.strip() for p in raw_pregs.split("|") if p.strip()]
        else:
            preguntas = [p.strip() for p in raw_pregs.split("\n") if p.strip()]

        for idx, preg in enumerate(preguntas):
            requests_batch.append(
                {
                    "createItem": {
                        "item": {
                            "title": preg,
                            "questionItem": {
                                "question": {
                                    "required": True,
                                    "textQuestion": {"paragraph": False},
                                }
                            },
                        },
                        "location": {"index": idx},
                    }
                }
            )
        if requests_batch:
            forms.forms().batchUpdate(
                formId=form_id, body={"requests": requests_batch}
            ).execute()
        responder_url = form_creado.get(
            "responderUri", f"https://docs.google.com/forms/d/{form_id}/viewform"
        )
        edit_url = f"https://docs.google.com/forms/d/{form_id}/edit"
        return (
            f"FORM_OK: Google Form creado exitosamente:\n"
            f"- Enlace para responder: {responder_url}\n"
            f"- Enlace para editar: {edit_url}"
        )
    except Exception as e:
        return format_google_error("Google Forms", e)


def planificar_ruta_google_maps(
    origen: str, destino: str, modo: str = "driving"
) -> str:
    """
    Calcula distancia, tiempo estimado y genera el enlace oficial de navegación en Google Maps.
    modo: 'driving' (auto), 'walking' (a pie) o 'transit' (transporte público).
    """
    log_tool(f"🗺️ Maps: {origen} -> {destino}")
    headers = {"User-Agent": "AgenteULimaODS4/1.0"}
    url_maps = (
        f"https://www.google.com/maps/dir/?api=1"
        f"&origin={quote_plus(origen)}&destination={quote_plus(destino)}&travelmode={modo}"
    )
    try:
        g1 = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": origen, "format": "json", "limit": 1},
            headers=headers,
            timeout=5,
        ).json()
        g2 = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": destino, "format": "json", "limit": 1},
            headers=headers,
            timeout=5,
        ).json()
        if g1 and g2:
            lon1, lat1 = g1[0]["lon"], g1[0]["lat"]
            lon2, lat2 = g2[0]["lon"], g2[0]["lat"]
            perfil = "foot" if modo == "walking" else "car"
            ruta = requests.get(
                f"https://router.project-osrm.org/route/v1/{perfil}/{lon1},{lat1};{lon2},{lat2}",
                params={"overview": "false"},
                timeout=5,
            ).json()
            if ruta.get("routes"):
                dist_km = round(ruta["routes"][0]["distance"] / 1000, 2)
                dur_min = round(ruta["routes"][0]["duration"] / 60)
                return (
                    f"Ruta calculada de **{origen}** a **{destino}**:\n"
                    f"- Distancia estimada: {dist_km} km\n"
                    f"- Tiempo aproximado ({modo}): {dur_min} min\n"
                    f"- Abrir navegación en Google Maps: {url_maps}"
                )
    except Exception:
        pass
    return f"Enlace directo de ruta en Google Maps de **{origen}** a **{destino}**: {url_maps}"


def gestionar_google_calendar(
    accion: str,
    titulo: str = "",
    fecha_inicio_iso: str = "",
    fecha_fin_iso: str = "",
    descripcion: str = "",
) -> str:
    """
    Lista próximos eventos o agenda un evento en Google Calendar.
    accion: 'listar' o 'agendar'.
    fecha_inicio_iso y fecha_fin_iso: formato 'YYYY-MM-DDTHH:MM:SS' (hora de Lima -05:00).
    """
    log_tool(f"📅 Calendar ({accion})")
    creds = get_user_credentials()
    if not creds:
        return "ERROR: Modo Rápido activo. Conecta tu cuenta de Google para usar Calendar."
    if not check_scope(creds, "https://www.googleapis.com/auth/calendar"):
        return "ERROR: Falta el permiso de Google Calendar. Cierra sesión y vuelve a entrar marcando todas las casillas."
    try:
        cal = build("calendar", "v3", credentials=creds)
        if accion.lower().strip() in ["agendar", "crear", "schedule", "event"] and titulo and fecha_inicio_iso:
            ini_clean = fecha_inicio_iso.strip()
            if len(ini_clean) == 10:
                ini_clean = f"{ini_clean}T09:00:00"
            if not fecha_fin_iso:
                dt_ini = datetime.datetime.fromisoformat(ini_clean)
                fecha_fin_iso = (dt_ini + datetime.timedelta(hours=1)).isoformat()
            else:
                fecha_fin_iso = fecha_fin_iso.strip()
                if len(fecha_fin_iso) == 10:
                    fecha_fin_iso = f"{fecha_fin_iso}T10:00:00"
            evento = {
                "summary": titulo,
                "description": descripcion,
                "start": {"dateTime": ini_clean, "timeZone": "America/Lima"},
                "end": {"dateTime": fecha_fin_iso, "timeZone": "America/Lima"},
            }
            creado = cal.events().insert(calendarId="primary", body=evento).execute()
            return f"EVENTO_OK: Evento agendado en Google Calendar: '{titulo}' -> {creado.get('htmlLink')}"
        ahora = datetime.datetime.now(datetime.timezone.utc).isoformat()
        eventos = (
            cal.events()
            .list(
                calendarId="primary",
                timeMin=ahora,
                maxResults=10,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
            .get("items", [])
        )
        if not eventos:
            return "No tienes eventos próximos en tu calendario."
        return "\n".join(
            [
                f"- {e.get('summary', 'Sin título')} ({e['start'].get('dateTime', e['start'].get('date'))})"
                for e in eventos
            ]
        )
    except Exception as e:
        return format_google_error("Google Calendar", e)


def gestionar_gmail(
    accion: str,
    destinatario: str = "",
    asunto: str = "",
    cuerpo: str = "",
    adjuntar_ultimo_archivo: bool = False,
) -> str:
    """
    Lee los últimos correos o envía un correo por Gmail (con opción de adjuntar el último archivo creado).
    accion: 'enviar' para mandar un correo, o 'leer' para revisar la bandeja.
    destinatario: dirección de correo electrónico válida (ej. 'usuario@gmail.com').
    """
    log_tool(f"📧 Gmail ({accion} -> {destinatario or 'bandeja'})")
    creds = get_user_credentials()
    if not creds:
        return "ERROR: Modo Rápido activo. Conecta tu cuenta de Google para usar Gmail."
    if not check_scope(creds, "https://www.googleapis.com/auth/gmail.modify"):
        return (
            "ERROR: Esta cuenta de Google no tiene activo el permiso de Gmail. "
            "Cierra sesión en la barra lateral, vuelve a entrar con Google y marca todas las casillas de permisos."
        )

    try:
        gmail = build("gmail", "v1", credentials=creds)
        if accion.lower().strip() in ["enviar", "send", "mandar", "enviar_correo", "write"]:
            dest_clean = destinatario.strip()
            if not dest_clean or not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", dest_clean):
                return (
                    "ERROR: Dirección de destinatario inválida o vacía. "
                    "Indica un correo válido (ej. nombre@gmail.com) o búscalo primero en Contactos."
                )
            msg = EmailMessage()
            msg["To"] = dest_clean
            msg["Subject"] = asunto.strip() if asunto else "Mensaje desde Agente IA ULima"
            msg.set_content(
                cuerpo.strip()
                if cuerpo
                else "Hola,\n\nEste mensaje fue enviado por el Agente IA de la Universidad de Lima."
            )
            ultimo = st.session_state.get("ultimo_archivo")
            adjuntado_txt = ""
            if adjuntar_ultimo_archivo and ultimo:
                msg.add_attachment(
                    ultimo["bytes"],
                    maintype="application",
                    subtype="octet-stream",
                    filename=ultimo["nombre"],
                )
                adjuntado_txt = f" con adjunto '{ultimo['nombre']}'"

            raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
            enviado = (
                gmail.users().messages().send(userId="me", body={"raw": raw}).execute()
            )
            remitente = st.session_state.get("user_email", "tu cuenta")
            return (
                f"ENVIADO_OK: Correo enviado desde {remitente} hacia {dest_clean}{adjuntado_txt} "
                f"(ID de mensaje Gmail: {enviado.get('id')})."
            )

        msgs = (
            gmail.users()
            .messages()
            .list(userId="me", maxResults=6)
            .execute()
            .get("messages", [])
        )
        if not msgs:
            return "Tu bandeja de entrada está vacía."
        resumenes = []
        for m in msgs:
            data = (
                gmail.users()
                .messages()
                .get(
                    userId="me",
                    id=m["id"],
                    format="metadata",
                    metadataHeaders=["From", "Subject"],
                )
                .execute()
            )
            headers = {
                h["name"]: h["value"] for h in data.get("payload", {}).get("headers", [])
            }
            resumenes.append(
                f"- De: {headers.get('From', '?')} | Asunto: {headers.get('Subject', '(Sin asunto)')} | Vista previa: {data.get('snippet', '')[:120]}"
            )
        return "\n".join(resumenes)
    except Exception as e:
        return format_google_error("Gmail", e)


def crear_documento_word_y_google_doc(titulo: str, contenido: str) -> str:
    """Crea un archivo Word (.docx) descargable y, si hay sesión Google, lo sube como Google Doc editable."""
    log_tool(f"📄 Word/Docs: {titulo}")
    doc = Document()
    doc.add_heading(titulo, level=1)
    for parrafo in contenido.split("\n"):
        if parrafo.strip():
            doc.add_paragraph(parrafo.strip())
    buf = io.BytesIO()
    doc.save(buf)
    data_bytes = buf.getvalue()
    nombre_archivo = f"{titulo[:40].strip().replace(' ', '_')}.docx"
    st.session_state["ultimo_archivo"] = {
        "nombre": nombre_archivo,
        "bytes": data_bytes,
    }
    st.session_state.setdefault("archivos_generados", []).append(
        st.session_state["ultimo_archivo"]
    )

    creds = get_user_credentials()
    if creds and check_scope(creds, "https://www.googleapis.com/auth/drive.file"):
        try:
            drive = build("drive", "v3", credentials=creds)
            media = MediaIoBaseUpload(
                io.BytesIO(data_bytes),
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            file_meta = {
                "name": titulo,
                "mimeType": "application/vnd.google-apps.document",
            }
            creado = (
                drive.files()
                .create(body=file_meta, media_body=media, fields="id,webViewLink")
                .execute()
            )
            return f"ARCHIVO_OK: Word generado ({nombre_archivo}) y subido a Google Docs: {creado.get('webViewLink')}"
        except Exception as e:
            return f"ARCHIVO_OK: Word generado para descarga local ({nombre_archivo}). Nota Drive: {e}"
    return f"ARCHIVO_OK: Documento Word '{nombre_archivo}' generado y listo para descargar en la barra lateral."


def crear_presentacion_ppt_y_google_slides(
    titulo_presentacion: str, diapositivas_texto: str
) -> str:
    """
    Crea un PowerPoint (.pptx) y lo sube a Google Slides.
    diapositivas_texto: separa cada diapositiva con '---' y usa la primera línea de cada bloque como título.
    """
    log_tool(f"📊 PPT/Slides: {titulo_presentacion}")
    prs = Presentation()
    sep = "---" if "---" in diapositivas_texto else "\n\n"
    bloques = [b.strip() for b in diapositivas_texto.split(sep) if b.strip()]
    if not bloques:
        bloques = [titulo_presentacion]
    for i, bloque in enumerate(bloques):
        lineas = [l.strip() for l in bloque.split("\n") if l.strip()]
        layout = prs.slide_layouts[0] if i == 0 else prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = lineas[0] if lineas else f"Diapositiva {i+1}"
        if len(lineas) > 1 and len(slide.placeholders) > 1:
            slide.placeholders[1].text = "\n".join(lineas[1:])
    buf = io.BytesIO()
    prs.save(buf)
    data_bytes = buf.getvalue()
    nombre_archivo = f"{titulo_presentacion[:40].strip().replace(' ', '_')}.pptx"
    st.session_state["ultimo_archivo"] = {
        "nombre": nombre_archivo,
        "bytes": data_bytes,
    }
    st.session_state.setdefault("archivos_generados", []).append(
        st.session_state["ultimo_archivo"]
    )

    creds = get_user_credentials()
    if creds and check_scope(creds, "https://www.googleapis.com/auth/drive.file"):
        try:
            drive = build("drive", "v3", credentials=creds)
            media = MediaIoBaseUpload(
                io.BytesIO(data_bytes),
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            file_meta = {
                "name": titulo_presentacion,
                "mimeType": "application/vnd.google-apps.presentation",
            }
            creado = (
                drive.files()
                .create(body=file_meta, media_body=media, fields="id,webViewLink")
                .execute()
            )
            return f"ARCHIVO_OK: PowerPoint generado ({nombre_archivo}) y subido a Google Slides: {creado.get('webViewLink')}"
        except Exception as e:
            return f"ARCHIVO_OK: PowerPoint generado para descarga local ({nombre_archivo}). Nota Drive: {e}"
    return f"ARCHIVO_OK: Presentación '{nombre_archivo}' lista para descargar en la barra lateral."


def crear_excel_y_google_sheets(nombre_hoja: str, filas_csv: str) -> str:
    """
    Crea un archivo Excel (.xlsx) y lo sube a Google Sheets.
    filas_csv: filas separadas por salto de línea y columnas separadas por '|'.
    """
    log_tool(f"📈 Excel/Sheets: {nombre_hoja}")
    wb = Workbook()
    ws = wb.active
    ws.title = nombre_hoja[:30]
    for linea in filas_csv.strip().split("\n"):
        if linea.strip():
            delim = "|" if "|" in linea else (";" if ";" in linea else ",")
            ws.append([celda.strip() for celda in linea.split(delim)])
    buf = io.BytesIO()
    wb.save(buf)
    data_bytes = buf.getvalue()
    nombre_archivo = f"{nombre_hoja[:40].strip().replace(' ', '_')}.xlsx"
    st.session_state["ultimo_archivo"] = {
        "nombre": nombre_archivo,
        "bytes": data_bytes,
    }
    st.session_state.setdefault("archivos_generados", []).append(
        st.session_state["ultimo_archivo"]
    )

    creds = get_user_credentials()
    if creds and check_scope(creds, "https://www.googleapis.com/auth/drive.file"):
        try:
            drive = build("drive", "v3", credentials=creds)
            media = MediaIoBaseUpload(
                io.BytesIO(data_bytes),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            file_meta = {
                "name": nombre_hoja,
                "mimeType": "application/vnd.google-apps.spreadsheet",
            }
            creado = (
                drive.files()
                .create(body=file_meta, media_body=media, fields="id,webViewLink")
                .execute()
            )
            return f"ARCHIVO_OK: Excel generado ({nombre_archivo}) y subido a Google Sheets: {creado.get('webViewLink')}"
        except Exception as e:
            return f"ARCHIVO_OK: Excel generado para descarga local ({nombre_archivo}). Nota Drive: {e}"
    return f"ARCHIVO_OK: Hoja de cálculo '{nombre_archivo}' lista para descargar en la barra lateral."


TOOLS_LIST = [
    buscar_internet,
    leer_pagina_web,
    buscar_videos_youtube,
    gestionar_google_tasks,
    buscar_contactos_google,
    crear_formulario_google,
    planificar_ruta_google_maps,
    gestionar_google_calendar,
    gestionar_gmail,
    crear_documento_word_y_google_doc,
    crear_presentacion_ppt_y_google_slides,
    crear_excel_y_google_sheets,
]


# ==============================================================================
# 3. MANEJO DE CALLBACK OAUTH 2.0
# ==============================================================================
if "code" in st.query_params and "google_creds" not in st.session_state:
    try:
        flow = get_oauth_flow()
        flow.fetch_token(code=st.query_params["code"])
        creds = flow.credentials
        st.session_state["google_creds"] = {
            "token": creds.token,
            "refresh_token": creds.refresh_token,
            "token_uri": creds.token_uri,
            "client_id": creds.client_id,
            "client_secret": creds.client_secret,
            "scopes": list(creds.scopes or []),
        }
        oauth2 = build("oauth2", "v2", credentials=creds)
        user_info = oauth2.userinfo().get().execute()
        st.session_state["user_email"] = user_info.get("email", "Usuario Google")
        st.session_state["modo_rapido"] = False
        st.query_params.clear()
        st.rerun()
    except Exception as e:
        st.query_params.clear()
        st.error(f"Error al conectar con Google: {e}")


# ==============================================================================
# 4. PANTALLA DE BIENVENIDA / LOGIN MULTI-TENANT
# ==============================================================================
if "google_creds" not in st.session_state and not st.session_state.get("modo_rapido"):
    _, col_centro, _ = st.columns([1, 2, 1])
    with col_centro:
        st.markdown("## 🎓 Agente IA — Universidad de Lima")
        st.write(
            "Conecta tu cuenta de Google para darle acceso aislado a tu **Gmail, Calendar, "
            "Docs, Slides, Sheets, YouTube, Tasks, Contactos, Forms y rutas de Maps**."
        )
        st.info(
            "💡 **Importante:** En la pantalla de permisos de Google, marca **"
            "Seleccionar todo** para habilitar el envío de correos, Calendar y Drive."
        )
        try:
            flow = get_oauth_flow()
            auth_url, _ = flow.authorization_url(
                prompt="consent", access_type="offline"
            )
            st.link_button(
                "🔵 Continuar con Google",
                auth_url,
                type="primary",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Revisa tus Secrets de OAuth en Streamlit: {e}")

        if st.button(
            "⚡ Entrar ahora en Modo Rápido (Web + Word + PPT + Excel + YouTube + Maps)",
            use_container_width=True,
        ):
            st.session_state["modo_rapido"] = True
            st.rerun()
    st.stop()


# ==============================================================================
# 5. INTERFAZ PRINCIPAL, VERIFICADOR DE PERMISOS Y BARRA LATERAL
# ==============================================================================
creds_actuales = get_user_credentials()
permisos_faltantes = []
if creds_actuales and creds_actuales.scopes:
    granted_set = set(creds_actuales.scopes)
    for sc, etiqueta in REQUIRED_WORKSPACE_SCOPES.items():
        if sc not in granted_set:
            permisos_faltantes.append(etiqueta)

with st.sidebar:
    st.markdown("## 🎓 Agente IA ULima")
    if st.session_state.get("user_email"):
        st.success(f"🟢 {st.session_state['user_email']}")
        if permisos_faltantes:
            st.warning(f"⚠️ Permisos sin marcar: {', '.join(permisos_faltantes)}")
        else:
            st.caption("✅ Todos los permisos Workspace activos")
    else:
        st.info("⚡ Modo Rápido (Sin cuenta Google)")

    st.divider()
    st.markdown("**Herramientas del Agente:**")
    st.markdown(
        """
- 🌐 Navegador Web & Lectura URL
- ▶️ YouTube Educativo (Búsqueda)
- ✅ Google Tasks (Pendientes)
- 👥 Google Contactos
- 📝 Google Forms (Encuestas/Quiz)
- 🗺️ Google Maps (Rutas y Distancias)
- 📅 Google Calendar (Leer y Agendar)
- 📧 Gmail (Leer y Enviar con adjuntos)
- 📄 Word + Google Docs
- 📊 PowerPoint + Google Slides
- 📈 Excel + Google Sheets
"""
    )

    if st.session_state.get("archivos_generados"):
        st.divider()
        st.markdown("**📥 Archivos descargables:**")
        for idx, arch in enumerate(st.session_state["archivos_generados"]):
            st.download_button(
                label=f"Descargar {arch['nombre']}",
                data=arch["bytes"],
                file_name=arch["nombre"],
                key=f"dl_{idx}_{arch['nombre']}",
                use_container_width=True,
            )

    st.divider()
    with st.expander("🔑 API Key de respaldo (opcional)"):
        st.text_input(
            "Pega otra GEMINI_API_KEY si cambias de cuenta:",
            type="password",
            key="custom_gemini_key",
        )

    if st.button("🗑️ Limpiar conversación", use_container_width=True):
        st.session_state["messages"] = []
        st.rerun()

    if st.button("🚪 Cerrar sesión", use_container_width=True):
        for k in ["google_creds", "user_email", "modo_rapido", "messages"]:
            st.session_state.pop(k, None)
        st.rerun()

if permisos_faltantes:
    st.warning(
        f"⚠️ Esta cuenta inició sesión sin autorizar: **{', '.join(permisos_faltantes)}**. "
        "Para que el agente pueda enviar correos o crear archivos en Drive/Calendar, "
        "pulsa el botón de abajo y marca **Seleccionar todo** en Google."
    )
    try:
        flow_reauth = get_oauth_flow()
        reauth_url, _ = flow_reauth.authorization_url(
            prompt="consent", access_type="offline"
        )
        col_re1, col_re2 = st.columns([1, 2])
        with col_re1:
            st.link_button(
                "🔓 Activar permisos faltantes",
                reauth_url,
                type="primary",
                use_container_width=True,
            )
    except Exception:
        pass

st.markdown("## ✨ ¿Qué quieres que haga hoy por ti?")
st.write(
    "Pídeme lo que quieras en lenguaje natural: investigar en internet o YouTube, "
    "orientarte entre las 14 carreras de la ULima, crear tareas, encuestas en Forms, "
    "calcular rutas en Maps, revisar/agendar Calendar, enviar Gmails o crear Docs, PPT y Excel."
)

c1, c2, c3, c4 = st.columns(4)
prompt_rapido = None
if c1.button("🧭 Orientación 14 Carreras ULima", use_container_width=True):
    prompt_rapido = "Ayúdame a comparar las 14 carreras de la Universidad de Lima según mis intereses."
if c2.button("▶️ Videos + Tareas en Tasks", use_container_width=True):
    prompt_rapido = "Busca 3 videos educativos en YouTube sobre Inteligencia Artificial en educación y créame una tarea en Google Tasks para estudiarlos esta semana."
if c3.button("📝 Crear Encuesta Google Forms", use_container_width=True):
    prompt_rapido = "Crea un Google Form titulado 'Encuesta de Hábitos de Estudio ULima' con 4 preguntas clave."
if c4.button("🗺️ Ruta a ULima en Maps", use_container_width=True):
    prompt_rapido = "Calcula la ruta y tiempo de viaje desde Miraflores, Lima hasta Universidad de Lima, Surco en Google Maps."

st.session_state.setdefault("messages", [])
for msg in st.session_state["messages"]:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        if msg.get("tools"):
            st.caption("🛠️ Acciones ejecutadas: " + " | ".join(msg["tools"]))

user_input = st.chat_input(
    "Escribe tu pedido aquí (ej. 'Busca videos de Cálculo, crea un resumen en Docs y agéndalo en Tasks')..."
)
prompt_final = prompt_rapido or user_input

if prompt_final:
    st.session_state["messages"].append({"role": "user", "content": prompt_final})
    with st.chat_message("user"):
        st.markdown(prompt_final)

    with st.chat_message("assistant"):
        with st.spinner("Ejecutando herramientas autónomas..."):
            st.session_state["tool_logs"] = []
            historial_gemini = [
                types.Content(
                    role="model" if m["role"] == "assistant" else "user",
                    parts=[types.Part.from_text(text=m["content"])],
                )
                for m in st.session_state["messages"][:-1]
            ]
            texto_respuesta = None
            errores = []

            for api_key in get_gemini_keys():
                if texto_respuesta:
                    break
                client = genai.Client(api_key=api_key)
                for modelo in MODELOS_DISPONIBLES:
                    for intento in range(2):
                        try:
                            st.session_state["tool_logs"] = []
                            chat = client.chats.create(
                                model=modelo,
                                history=historial_gemini,
                                config=types.GenerateContentConfig(
                                    system_instruction=get_system_instruction(),
                                    tools=TOOLS_LIST,
                                    temperature=0.3,
                                    automatic_function_calling=types.AutomaticFunctionCallingConfig(
                                        maximum_remote_calls=6
                                    ),
                                ),
                            )
                            response = chat.send_message(prompt_final)
                            texto_respuesta = response.text or "Acción completada."
                            break
                        except Exception as e:
                            err_str = str(e)
                            if (
                                "503" in err_str or "UNAVAILABLE" in err_str
                            ) and intento == 0:
                                time.sleep(1.5)
                                continue
                            errores.append(f"{modelo}: {err_str}")
                            break
                    if texto_respuesta:
                        break

            if not texto_respuesta:
                texto_respuesta = (
                    "⚠️ No se pudo completar con los modelos disponibles:\n- "
                    + "\n- ".join(errores)
                )

            herramientas_usadas = list(
                dict.fromkeys(st.session_state.get("tool_logs", []))
            )
            st.markdown(texto_respuesta)
            if herramientas_usadas:
                st.caption("🛠️ Acciones ejecutadas: " + " | ".join(herramientas_usadas))

            st.session_state["messages"].append(
                {
                    "role": "assistant",
                    "content": texto_respuesta,
                    "tools": herramientas_usadas,
                }
            )
            st.rerun()
