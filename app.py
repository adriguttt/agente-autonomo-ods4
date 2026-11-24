import os

os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"] = "1"

import io
import base64
import datetime
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

SYSTEM_INSTRUCTION = """
Eres el Agente IA Autónomo de la Universidad de Lima enfocado en el ODS 4 (Educación de Calidad).
Ayudas a estudiantes, docentes y postulantes a investigar, estudiar, gestionar su tiempo y automatizar tareas reales.

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

Reglas de ejecución:
- Ejecuta las herramientas disponibles de forma autónoma y encadenada cuando el usuario pida varias acciones en un mismo mensaje.
- Si el usuario pide enviar un correo a alguien por su nombre (ej. "mándale un correo a Carlos"), busca primero su correo en Google Contactos antes de enviar.
- Si el usuario pide crear un documento, presentación u hoja de cálculo y enviarlo por Gmail, genera primero el archivo y luego usa gestionar_gmail con adjuntar_ultimo_archivo=True.
- Responde siempre en español claro, estructurado y con enlaces directos a los recursos creados (Docs, Slides, Sheets, Forms, YouTube o Google Maps).
- Fecha y hora actual de referencia: """ + datetime.datetime.now().strftime("%Y-%m-%d %H:%M")


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


# ==============================================================================
# 2. HERRAMIENTAS AUTÓNOMAS DEL AGENTE (AISLADAS POR SESIÓN)
# ==============================================================================
def buscar_internet(consulta: str) -> str:
    """Busca información actualizada en internet y devuelve títulos, enlaces y resúmenes."""
    try:
        resultados = []
        with DDGS() as ddgs:
            for r in ddgs.text(consulta, max_results=5):
                resultados.append(
                    f"- {r.get('title')}: {r.get('body')} (URL: {r.get('href')})"
                )
        return "\n".join(resultados) if resultados else "No se encontraron resultados."
    except Exception as e:
        return f"Error en búsqueda web: {e}"


def leer_pagina_web(url: str) -> str:
    """Extrae el texto principal de una URL pública para analizar artículos o documentos web."""
    try:
        resp = requests.get(url, timeout=12, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer"]):
            tag.decompose()
        texto = " ".join(soup.stripped_strings)
        return texto[:6000] if texto else "La página no contiene texto legible."
    except Exception as e:
        return f"No se pudo leer la URL: {e}"


def buscar_videos_youtube(tema: str, max_resultados: int = 5) -> str:
    """Busca videos educativos, tutoriales o conferencias en YouTube sobre un tema académico."""
    limite = min(max_resultados, 8)
    try:
        yt = build("youtube", "v3", developerKey=st.secrets["GEMINI_API_KEY"])
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
        with DDGS() as ddgs:
            for r in ddgs.text(f"site:youtube.com/watch {tema}", max_results=limite):
                resultados.append(
                    f"- **{r.get('title')}** -> {r.get('href')}"
                )
        return (
            "\n".join(resultados)
            if resultados
            else "No se encontraron videos en YouTube para ese tema."
        )
    except Exception as e:
        return f"Error al buscar videos en YouTube: {e}"


def gestionar_google_tasks(
    accion: str, titulo: str = "", notas: str = "", fecha_limite: str = ""
) -> str:
    """
    Lista o crea tareas en Google Tasks.
    accion: 'listar' o 'crear'.
    fecha_limite: opcional en formato 'YYYY-MM-DD'.
    """
    creds = get_user_credentials()
    if not creds:
        return "Modo Rápido activo: conecta tu cuenta de Google para usar Google Tasks."
    try:
        service = build("tasks", "v1", credentials=creds)
        if accion.lower() == "crear" and titulo:
            body = {"title": titulo, "notes": notas}
            if fecha_limite:
                body["due"] = f"{fecha_limite}T23:59:00.000Z"
            creada = service.tasks().insert(tasklist="@default", body=body).execute()
            return f"Tarea creada en Google Tasks: '{creada.get('title')}' (ID: {creada.get('id')})."
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
        return f"Error en Google Tasks: {e}"


def buscar_contactos_google(nombre_o_correo: str) -> str:
    """Busca personas en Google Contactos del usuario para obtener sus correos o teléfonos."""
    creds = get_user_credentials()
    if not creds:
        return "Modo Rápido activo: conecta tu cuenta de Google para buscar en Contactos."
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
        return f"Error al buscar en Contactos: {e}"


def crear_formulario_google(
    titulo: str, descripcion: str, preguntas_separadas_por_punto_y_coma: str
) -> str:
    """
    Crea un formulario o encuesta real en Google Forms con preguntas de respuesta corta.
    preguntas_separadas_por_punto_y_coma: ejemplo '¿Cuál es tu carrera?;¿Qué tema prefieres?'
    """
    creds = get_user_credentials()
    if not creds:
        return "Modo Rápido activo: conecta tu cuenta de Google para crear Google Forms."
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
        preguntas = [
            p.strip()
            for p in preguntas_separadas_por_punto_y_coma.split(";")
            if p.strip()
        ]
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
            f"Google Form creado exitosamente:\n"
            f"- Enlace para responder: {responder_url}\n"
            f"- Enlace para editar: {edit_url}"
        )
    except Exception as e:
        return f"Error al crear Google Form: {e}"


def planificar_ruta_google_maps(
    origen: str, destino: str, modo: str = "driving"
) -> str:
    """
    Calcula distancia, tiempo estimado y genera el enlace oficial de navegación en Google Maps.
    modo: 'driving' (auto), 'walking' (a pie) o 'transit' (transporte público).
    """
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
            timeout=8,
        ).json()
        g2 = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": destino, "format": "json", "limit": 1},
            headers=headers,
            timeout=8,
        ).json()
        if g1 and g2:
            lon1, lat1 = g1[0]["lon"], g1[0]["lat"]
            lon2, lat2 = g2[0]["lon"], g2[0]["lat"]
            perfil = "foot" if modo == "walking" else "car"
            ruta = requests.get(
                f"https://router.project-osrm.org/route/v1/{perfil}/{lon1},{lat1};{lon2},{lat2}",
                params={"overview": "false"},
                timeout=8,
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
    creds = get_user_credentials()
    if not creds:
        return "Modo Rápido activo: conecta tu cuenta de Google para usar Calendar."
    try:
        cal = build("calendar", "v3", credentials=creds)
        if accion.lower() == "agendar" and titulo and fecha_inicio_iso:
            if not fecha_fin_iso:
                dt_ini = datetime.datetime.fromisoformat(fecha_inicio_iso)
                fecha_fin_iso = (dt_ini + datetime.timedelta(hours=1)).isoformat()
            evento = {
                "summary": titulo,
                "description": descripcion,
                "start": {"dateTime": fecha_inicio_iso, "timeZone": "America/Lima"},
                "end": {"dateTime": fecha_fin_iso, "timeZone": "America/Lima"},
            }
            creado = cal.events().insert(calendarId="primary", body=evento).execute()
            return f"Evento agendado: '{titulo}' -> {creado.get('htmlLink')}"
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
        return f"Error en Google Calendar: {e}"


def gestionar_gmail(
    accion: str,
    destinatario: str = "",
    asunto: str = "",
    cuerpo: str = "",
    adjuntar_ultimo_archivo: bool = False,
) -> str:
    """
    Lee los últimos correos o envía un correo por Gmail (con opción de adjuntar el último archivo creado).
    accion: 'leer' o 'enviar'.
    """
    creds = get_user_credentials()
    if not creds:
        return "Modo Rápido activo: conecta tu cuenta de Google para usar Gmail."
    try:
        gmail = build("gmail", "v1", credentials=creds)
        if accion.lower() == "enviar" and destinatario:
            msg = EmailMessage()
            msg["To"] = destinatario
            msg["Subject"] = asunto or "Mensaje desde Agente IA ULima"
            msg.set_content(cuerpo or "")
            ultimo = st.session_state.get("ultimo_archivo")
            if adjuntar_ultimo_archivo and ultimo:
                msg.add_attachment(
                    ultimo["bytes"],
                    maintype="application",
                    subtype="octet-stream",
                    filename=ultimo["nombre"],
                )
            raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
            gmail.users().messages().send(userId="me", body={"raw": raw}).execute()
            return f"Correo enviado correctamente a {destinatario}."
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
        return f"Error en Gmail: {e}"


def crear_documento_word_y_google_doc(titulo: str, contenido: str) -> str:
    """Crea un archivo Word (.docx) descargable y, si hay sesión Google, lo sube como Google Doc editable."""
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
    if creds:
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
            return f"Word generado ({nombre_archivo}) y subido a Google Docs: {creado.get('webViewLink')}"
        except Exception as e:
            return f"Word generado para descarga ({nombre_archivo}). Nota Drive: {e}"
    return f"Documento Word '{nombre_archivo}' generado y listo para descargar en la barra lateral."


def crear_presentacion_ppt_y_google_slides(
    titulo_presentacion: str, diapositivas_texto: str
) -> str:
    """
    Crea un PowerPoint (.pptx) y lo sube a Google Slides.
    diapositivas_texto: separa cada diapositiva con '---' y usa la primera línea de cada bloque como título.
    """
    prs = Presentation()
    bloques = [b.strip() for b in diapositivas_texto.split("---") if b.strip()]
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
    if creds:
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
            return f"PowerPoint generado ({nombre_archivo}) y subido a Google Slides: {creado.get('webViewLink')}"
        except Exception as e:
            return f"PowerPoint generado ({nombre_archivo}). Nota Drive: {e}"
    return f"Presentación '{nombre_archivo}' lista para descargar en la barra lateral."


def crear_excel_y_google_sheets(nombre_hoja: str, filas_csv: str) -> str:
    """
    Crea un archivo Excel (.xlsx) y lo sube a Google Sheets.
    filas_csv: filas separadas por salto de línea y columnas separadas por '|'.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = nombre_hoja[:30]
    for linea in filas_csv.strip().split("\n"):
        if linea.strip():
            ws.append([celda.strip() for celda in linea.split("|")])
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
    if creds:
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
            return f"Excel generado ({nombre_archivo}) y subido a Google Sheets: {creado.get('webViewLink')}"
        except Exception as e:
            return f"Excel generado ({nombre_archivo}). Nota Drive: {e}"
    return f"Hoja de cálculo '{nombre_archivo}' lista para descargar en la barra lateral."


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
            "scopes": list(creds.scopes or SCOPES),
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
# 5. INTERFAZ PRINCIPAL Y BARRA LATERAL
# ==============================================================================
with st.sidebar:
    st.markdown("## 🎓 Agente IA ULima")
    if st.session_state.get("user_email"):
        st.success(f"🟢 {st.session_state['user_email']}")
        st.caption("Conectado a Workspace, YouTube, Tasks, Forms y Maps")
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
    if st.button("🗑️ Limpiar conversación", use_container_width=True):
        st.session_state["messages"] = []
        st.rerun()

    if st.button("🚪 Cerrar sesión", use_container_width=True):
        for k in ["google_creds", "user_email", "modo_rapido", "messages"]:
            st.session_state.pop(k, None)
        st.rerun()

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
            try:
                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                historial_gemini = [
                    types.Content(
                        role="model" if m["role"] == "assistant" else "user",
                        parts=[types.Part.from_text(text=m["content"])],
                    )
                    for m in st.session_state["messages"][:-1]
                ]
                chat = client.chats.create(
                    model="gemini-2.5-flash",
                    history=historial_gemini,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        tools=TOOLS_LIST,
                        temperature=0.4,
                    ),
                )
                response = chat.send_message(prompt_final)
                texto_respuesta = response.text or "Acción completada."
            except Exception as e:
                texto_respuesta = f"Ocurrió un error al procesar la solicitud: {e}"

            st.markdown(texto_respuesta)
            st.session_state["messages"].append(
                {"role": "assistant", "content": texto_respuesta}
            )
