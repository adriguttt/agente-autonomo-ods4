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
import streamlit.components.v1 as components
from bs4 import BeautifulSoup
from duckduckgo_search import DDGS
from docx import Document
from pptx import Presentation
from openpyxl import Workbook, load_workbook
from pypdf import PdfReader
from PIL import Image
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

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
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/tasks",
    "https://www.googleapis.com/auth/contacts.readonly",
    "https://www.googleapis.com/auth/forms.body",
]

REQUIRED_WORKSPACE_SCOPES = {
    "https://www.googleapis.com/auth/gmail.modify": "Gmail",
    "https://www.googleapis.com/auth/calendar": "Calendar",
    "https://www.googleapis.com/auth/drive.file": "Drive (Crear/Subir)",
    "https://www.googleapis.com/auth/drive.readonly": "Drive (Leer archivos)",
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


def registrar_archivo_generado(nombre: str, data_bytes: bytes, mime: str = "application/octet-stream") -> None:
    item = {"nombre": nombre, "bytes": data_bytes, "mime": mime}
    st.session_state["ultimo_archivo"] = item
    st.session_state.setdefault("archivos_generados", []).append(item)


def get_system_instruction() -> str:
    ahora_lima = datetime.datetime.now(ZoneInfo("America/Lima")).strftime(
        "%Y-%m-%d %H:%M (%A)"
    )
    usuario = st.session_state.get("user_email", "Modo Rápido (sin cuenta Google)")
    adjuntos = st.session_state.get("archivos_subidos_info", [])
    resumen_adj = ", ".join(adjuntos) if adjuntos else "Ninguno"
    return f"""
Eres el Agente IA Autónomo de la Universidad de Lima enfocado en el ODS 4 (Educación de Calidad).
Ayudas a estudiantes, docentes y postulantes a investigar, estudiar, analizar documentos/imágenes, construir aplicaciones interactivas (Canvas) y automatizar tareas reales.
Usuario conectado actualmente: {usuario}.
Archivos/imágenes adjuntados por el usuario en sesión: {resumen_adj}.

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
- JAMÁS afirmes que enviaste un correo, creaste un archivo, agendaste un evento o creaste una tarea/encuesta/canvas si la herramienta no devolvió una confirmación explícita ('ENVIADO_OK', 'ARCHIVO_OK', 'EVENTO_OK', 'TAREA_OK', 'FORM_OK', 'CANVAS_OK', 'IMAGEN_OK') o si devolvió 'ERROR:'. Si hay un error, explica el motivo exacto al usuario.
- Cuando el usuario pida crear una app interactiva, simulador, calculadora, mapa mental visual, diagrama o página web (Canvas), llama a `crear_canvas_interactivo` entregando código HTML5 + CSS + JS autocontenido y funcional.
- Cuando el usuario pida generar una imagen, gráfico de barras/líneas/pastel o esquema visual descargable, usa `generar_imagen_o_grafico`.
- Cuando el usuario pida buscar o leer archivos de su Google Drive, usa `gestionar_google_drive`.
- Para enviar correos usa `gestionar_gmail` con accion='enviar' y un correo electrónico válido en destinatario. Si solo te dan un nombre, consulta primero `buscar_contactos_google`.
- Si el usuario pide crear un documento, presentación, Excel, imagen o Canvas y enviarlo por Gmail, genéralo primero y luego llama a `gestionar_gmail` con adjuntar_ultimo_archivo=True.
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
                f"ERROR: Permiso denegado en {servicio} (403). Cierra sesión, vuelve a entrar "
                f"marcando todas las casillas de permisos y verifica que la API de {servicio} esté activa."
            )
        if status == 401:
            return "ERROR: La sesión de Google expiró (401). Cierra sesión y vuelve a entrar."
    return f"ERROR en {servicio}: {err}"


def extraer_texto_archivo(nombre: str, data_bytes: bytes) -> str:
    ext = nombre.lower().split(".")[-1]
    try:
        if ext == "pdf":
            reader = PdfReader(io.BytesIO(data_bytes))
            paginas = [p.extract_text() or "" for p in reader.pages[:25]]
            return "\n".join(paginas).strip()[:12000]
        if ext == "docx":
            doc = Document(io.BytesIO(data_bytes))
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])[:12000]
        if ext == "xlsx":
            wb = load_workbook(io.BytesIO(data_bytes), data_only=True)
            salida = []
            for sheet in wb.sheetnames[:5]:
                ws = wb[sheet]
                salida.append(f"--- Hoja: {sheet} ---")
                for row in list(ws.iter_rows(values_only=True))[:60]:
                    if any(row):
                        salida.append(" | ".join([str(c) if c is not None else "" for c in row]))
            return "\n".join(salida)[:12000]
        if ext == "pptx":
            prs = Presentation(io.BytesIO(data_bytes))
            salida = []
            for i, slide in enumerate(prs.slides[:30]):
                textos = [sh.text.strip() for sh in slide.shapes if hasattr(sh, "text") and sh.text.strip()]
                salida.append(f"[Diapositiva {i+1}]: " + " / ".join(textos))
            return "\n".join(salida)[:12000]
        return data_bytes.decode("utf-8", errors="ignore")[:12000]
    except Exception as e:
        return f"(No se pudo extraer texto de {nombre}: {e})"


# ==============================================================================
# 2. HERRAMIENTAS AUTÓNOMAS DEL AGENTE (15 HERRAMIENTAS)
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


def gestionar_google_drive(accion: str, consulta_o_id: str = "") -> str:
    """
    Busca archivos en el Google Drive del usuario o lee el contenido de un archivo por su ID.
    accion: 'buscar' (lista archivos que coincidan con consulta_o_id) o 'leer' (lee el texto del archivo cuyo file_id es consulta_o_id).
    """
    log_tool(f"📁 Drive ({accion}: {consulta_o_id[:25]})")
    creds = get_user_credentials()
    if not creds:
        return "ERROR: Modo Rápido activo. Conecta tu cuenta de Google para acceder a Google Drive."
    if not (
        check_scope(creds, "https://www.googleapis.com/auth/drive.readonly")
        or check_scope(creds, "https://www.googleapis.com/auth/drive.file")
    ):
        return "ERROR: Falta el permiso de lectura de Google Drive. Pulsa 'Activar permisos faltantes' arriba."
    try:
        drive = build("drive", "v3", credentials=creds)
        if accion.lower().strip() in ["leer", "read", "abrir", "open"] and consulta_o_id:
            fid = consulta_o_id.strip()
            meta = drive.files().get(fileId=fid, fields="id,name,mimeType").execute()
            mime = meta.get("mimeType", "")
            nombre = meta.get("name", "archivo")
            if mime == "application/vnd.google-apps.document":
                contenido = drive.files().export(fileId=fid, mimeType="text/plain").execute()
                return f"Contenido de Google Doc '{nombre}':\n{contenido.decode('utf-8', errors='ignore')[:10000]}"
            if mime == "application/vnd.google-apps.spreadsheet":
                contenido = drive.files().export(fileId=fid, mimeType="text/csv").execute()
                return f"Contenido de Google Sheet '{nombre}':\n{contenido.decode('utf-8', errors='ignore')[:10000]}"
            if mime == "application/vnd.google-apps.presentation":
                contenido = drive.files().export(fileId=fid, mimeType="text/plain").execute()
                return f"Contenido de Google Slides '{nombre}':\n{contenido.decode('utf-8', errors='ignore')[:10000]}"
            raw_bytes = drive.files().get_media(fileId=fid).execute()
            texto = extraer_texto_archivo(nombre, raw_bytes)
            return f"Contenido de '{nombre}':\n{texto[:10000]}"

        q = "trashed = false"
        if consulta_o_id.strip():
            termino = consulta_o_id.strip().replace("'", "\\'")
            q += f" and name contains '{termino}'"
        res = (
            drive.files()
            .list(
                q=q,
                pageSize=12,
                orderBy="modifiedTime desc",
                fields="files(id, name, mimeType, webViewLink, modifiedTime)",
            )
            .execute()
        )
        archivos = res.get("files", [])
        if not archivos:
            return f"No se encontraron archivos en Google Drive con '{consulta_o_id}'."
        return "\n".join(
            [
                f"- **{f['name']}** | ID: `{f['id']}` | Tipo: {f['mimeType'].split('.')[-1]} | Enlace: {f.get('webViewLink', '')}"
                for f in archivos
            ]
        )
    except Exception as e:
        return format_google_error("Google Drive", e)


def crear_canvas_interactivo(titulo: str, codigo_html_completo: str) -> str:
    """
    Crea y renderiza en pantalla un Canvas interactivo (aplicación web, simulador, calculadora,
    diagrama visual o dashboard) en HTML5/CSS/JS autocontenido y lo guarda como archivo descargable.
    """
    log_tool(f"🎨 Canvas: {titulo}")
    html_limpio = codigo_html_completo.strip()
    if "<html" not in html_limpio.lower():
        html_limpio = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{titulo}</title></head>
<body style="font-family: system-ui, sans-serif; margin: 0; padding: 16px;">
{html_limpio}
</body></html>"""

    nombre_archivo = f"Canvas_{titulo[:35].strip().replace(' ', '_')}.html"
    data_bytes = html_limpio.encode("utf-8")
    registrar_archivo_generado(nombre_archivo, data_bytes, "text/html")
    st.session_state["canvas_activo"] = {
        "titulo": titulo,
        "html": html_limpio,
        "nombre_archivo": nombre_archivo,
    }
    return (
        f"CANVAS_OK: Canvas interactivo '{titulo}' renderizado en pantalla y guardado como '{nombre_archivo}' "
        "para descarga o envío por Gmail."
    )


def generar_imagen_o_grafico(
    titulo: str,
    tipo: str,
    etiquetas_separadas_por_coma: str,
    valores_separados_por_coma: str,
    subtitulo_o_nota: str = "",
) -> str:
    """
    Genera una imagen PNG de alta calidad (gráfico estadístico de barras, líneas, pastel o tarjeta visual académica)
    descargable y adjuntable por correo.
    tipo: 'barras', 'lineas', 'pastel' o 'tarjeta'.
    """
    log_tool(f"🖼️ Imagen/Gráfico: {titulo}")
    try:
        fig, ax = plt.subplots(figsize=(8, 4.5), dpi=160)
        fig.patch.set_facecolor("#FFFFFF")
        t_clean = tipo.lower().strip()

        if t_clean in ["barras", "lineas", "pastel"] and etiquetas_separadas_por_coma and valores_separados_por_coma:
            etiquetas = [e.strip() for e in etiquetas_separadas_por_coma.split(",") if e.strip()]
            valores = []
            for v in valores_separados_por_coma.split(","):
                v_num = re.sub(r"[^\d.\-]", "", v.strip())
                if v_num:
                    valores.append(float(v_num))
            n = min(len(etiquetas), len(valores))
            etiquetas, valores = etiquetas[:n], valores[:n]
            colores = ["#FF6B00", "#1E293B", "#0EA5E9", "#10B981", "#8B5CF6", "#F59E0B"]

            if t_clean == "pastel":
                ax.pie(valores, labels=etiquetas, autopct="%1.1f%%", colors=colores[:n], startangle=140)
            elif t_clean == "lineas":
                ax.plot(etiquetas, valores, marker="o", linewidth=2.5, color="#FF6B00")
                ax.grid(True, linestyle="--", alpha=0.4)
                plt.xticks(rotation=20, ha="right")
            else:
                bars = ax.bar(etiquetas, valores, color=colores[:n])
                ax.bar_label(bars, padding=3, fontsize=9)
                ax.grid(axis="y", linestyle="--", alpha=0.4)
                plt.xticks(rotation=20, ha="right")
            ax.set_title(titulo, fontsize=13, fontweight="bold", pad=12, color="#0F172A")
            if subtitulo_o_nota:
                fig.text(0.5, 0.01, subtitulo_o_nota, ha="center", fontsize=8.5, color="#475569")
        else:
            ax.axis("off")
            fig.patch.set_facecolor("#0F172A")
            ax.text(0.5, 0.72, "UNIVERSIDAD DE LIMA — ODS 4", ha="center", va="center", fontsize=10, color="#FF8C38", fontweight="bold")
            ax.text(0.5, 0.50, titulo, ha="center", va="center", fontsize=15, color="#FFFFFF", fontweight="bold", wrap=True)
            ax.text(0.5, 0.25, subtitulo_o_nota or etiquetas_separadas_por_coma, ha="center", va="center", fontsize=10.5, color="#CBD5E1", wrap=True)

        plt.tight_layout()
        buf = io.BytesIO()
        plt.savefig(buf, format="png", bbox_inches="tight")
        plt.close(fig)
        img_bytes = buf.getvalue()

        nombre_archivo = f"Imagen_{titulo[:35].strip().replace(' ', '_')}.png"
        registrar_archivo_generado(nombre_archivo, img_bytes, "image/png")
        st.session_state["ultima_imagen"] = {"titulo": titulo, "bytes": img_bytes, "nombre": nombre_archivo}
        return f"IMAGEN_OK: Imagen '{nombre_archivo}' generada, mostrada en pantalla y lista para descargar o enviar por Gmail."
    except Exception as e:
        return f"ERROR al generar imagen: {e}"


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
    Lee los últimos correos o envía un correo por Gmail (con opción de adjuntar el último archivo creado o subido).
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
            "Pulsa 'Activar permisos faltantes' arriba y marca todas las casillas."
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
    registrar_archivo_generado(
        nombre_archivo,
        data_bytes,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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
    registrar_archivo_generado(
        nombre_archivo,
        data_bytes,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
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
    registrar_archivo_generado(
        nombre_archivo,
        data_bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
    gestionar_google_drive,
    crear_canvas_interactivo,
    generar_imagen_o_grafico,
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
            "Google Drive, Docs, Slides, Sheets, YouTube, Tasks, Contactos, Forms, Maps y Canvas**."
        )
        st.info(
            "💡 **Importante:** En la pantalla de permisos de Google, marca **"
            "Seleccionar todo** para habilitar Gmail, Calendar y Google Drive."
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
            "⚡ Entrar ahora en Modo Rápido (Canvas + Adjuntos + Web + Office + YouTube)",
            use_container_width=True,
        ):
            st.session_state["modo_rapido"] = True
            st.rerun()
    st.stop()


# ==============================================================================
# 5. INTERFAZ PRINCIPAL, ADJUNTOS, VERIFICADOR DE PERMISOS Y BARRA LATERAL
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
    st.markdown("**📎 Adjuntar archivos o imágenes:**")
    archivos_subidos = st.file_uploader(
        "Sube PDF, Word, Excel, PPT, TXT o imágenes (PNG/JPG):",
        type=["pdf", "docx", "xlsx", "pptx", "txt", "csv", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key="uploader_sidebar",
    )

    contexto_documentos_subidos = []
    imagenes_subidas_parts = []
    nombres_subidos = []

    if archivos_subidos:
        for f in archivos_subidos:
            f_bytes = f.getvalue()
            nombres_subidos.append(f.name)
            if f.type and f.type.startswith("image/"):
                imagenes_subidas_parts.append(
                    types.Part.from_bytes(data=f_bytes, mime_type=f.type)
                )
                st.image(f_bytes, caption=f.name, use_container_width=True)
            else:
                texto_doc = extraer_texto_archivo(f.name, f_bytes)
                contexto_documentos_subidos.append(
                    f"=== CONTENIDO DEL ARCHIVO ADJUNTO '{f.name}' ===\n{texto_doc}"
                )
        # Guarda el último archivo subido como disponible para reenviar por Gmail
        ultimo_subido = archivos_subidos[-1]
        st.session_state["ultimo_archivo"] = {
            "nombre": ultimo_subido.name,
            "bytes": ultimo_subido.getvalue(),
            "mime": ultimo_subido.type or "application/octet-stream",
        }
    st.session_state["archivos_subidos_info"] = nombres_subidos

    st.divider()
    st.markdown("**Herramientas del Agente (15):**")
    st.markdown(
        """
- 🎨 Canvas Interactivo (Apps/Web/Diagramas)
- 🖼️ Generador de Imágenes y Gráficos
- 📎 Análisis de PDF, Office e Imágenes
- 📁 Google Drive (Buscar y Leer archivos)
- 🌐 Navegador Web & Lectura URL
- ▶️ YouTube Educativo
- ✅ Google Tasks & 📅 Calendar
- 📧 Gmail (Enviar con adjuntos) & 👥 Contactos
- 📝 Google Forms & 🗺️ Google Maps
- 📄 Word/Docs | 📊 PPT/Slides | 📈 Excel/Sheets
"""
    )

    if st.session_state.get("archivos_generados"):
        st.divider()
        st.markdown("**📥 Archivos creados por el agente:**")
        for idx, arch in enumerate(st.session_state["archivos_generados"]):
            st.download_button(
                label=f"Descargar {arch['nombre']}",
                data=arch["bytes"],
                file_name=arch["nombre"],
                mime=arch.get("mime", "application/octet-stream"),
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

    if st.button("🗑️ Limpiar conversación y Canvas", use_container_width=True):
        st.session_state["messages"] = []
        st.session_state.pop("canvas_activo", None)
        st.session_state.pop("ultima_imagen", None)
        st.rerun()

    if st.button("🚪 Cerrar sesión", use_container_width=True):
        for k in ["google_creds", "user_email", "modo_rapido", "messages", "canvas_activo", "ultima_imagen"]:
            st.session_state.pop(k, None)
        st.rerun()

if permisos_faltantes:
    st.warning(
        f"⚠️ Esta cuenta inició sesión sin autorizar: **{', '.join(permisos_faltantes)}**. "
        "Pulsa el botón de abajo y marca **Seleccionar todo** en Google para activar Google Drive y Gmail."
    )
    try:
        flow_reauth = get_oauth_flow()
        reauth_url, _ = flow_reauth.authorization_url(
            prompt="consent", access_type="offline"
        )
        col_re1, _ = st.columns([1, 2])
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
    "Adjunta documentos o imágenes en la barra lateral, pídeme que revise tu **Google Drive**, "
    "cree un **Canvas interactivo**, genere **gráficos/imágenes**, agende en **Calendar** o envíe **Gmails**."
)

c1, c2, c3, c4 = st.columns(4)
prompt_rapido = None
if c1.button("🎨 Crear Canvas Simulador ULima", use_container_width=True):
    prompt_rapido = "Crea un Canvas interactivo con una calculadora de promedio ponderado y simulador de notas para estudiantes de la Universidad de Lima."
if c2.button("📁 Revisar mi Google Drive", use_container_width=True):
    prompt_rapido = "Busca mis últimos archivos en Google Drive y muéstrame su lista con enlaces."
if c3.button("🖼️ Generar Gráfico ODS 4", use_container_width=True):
    prompt_rapido = "Genera una imagen con un gráfico de barras sobre el impacto de la IA en la educación superior (Comprensión 85, Productividad 92, Investigación 88, Redacción 79)."
if c4.button("🧭 Orientación 14 Carreras ULima", use_container_width=True):
    prompt_rapido = "Ayúdame a comparar las 14 carreras de la Universidad de Lima según mis intereses."

# Renderizado persistente del Canvas activo si existe
if st.session_state.get("canvas_activo"):
    canvas = st.session_state["canvas_activo"]
    with st.expander(f"🎨 Canvas Interactivo: {canvas['titulo']}", expanded=True):
        components.html(canvas["html"], height=520, scrolling=True)
        st.download_button(
            label=f"⬇️ Descargar {canvas['nombre_archivo']}",
            data=canvas["html"].encode("utf-8"),
            file_name=canvas["nombre_archivo"],
            mime="text/html",
            key="dl_canvas_main",
        )

# Renderizado persistente de la última imagen generada si existe
if st.session_state.get("ultima_imagen"):
    img_data = st.session_state["ultima_imagen"]
    with st.expander(f"🖼️ Imagen Generada: {img_data['titulo']}", expanded=True):
        st.image(img_data["bytes"], caption=img_data["titulo"], use_container_width=True)
        st.download_button(
            label=f"⬇️ Descargar {img_data['nombre']}",
            data=img_data["bytes"],
            file_name=img_data["nombre"],
            mime="image/png",
            key="dl_img_main",
        )

st.session_state.setdefault("messages", [])
for msg in st.session_state["messages"]:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        if msg.get("tools"):
            st.caption("🛠️ Acciones ejecutadas: " + " | ".join(msg["tools"]))

user_input = st.chat_input(
    "Escribe tu pedido aquí (ej. 'Analiza el PDF adjunto, crea un Canvas resumen y envíalo por Gmail')..."
)
prompt_final = prompt_rapido or user_input

if prompt_final:
    etiqueta_usuario = prompt_final
    if nombres_subidos:
        etiqueta_usuario += f"\n\n📎 *Archivos adjuntos activos: {', '.join(nombres_subidos)}*"

    st.session_state["messages"].append({"role": "user", "content": etiqueta_usuario})
    with st.chat_message("user"):
        st.markdown(etiqueta_usuario)

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

            # Construir partes del mensaje actual (texto + documentos extraídos + imágenes)
            partes_mensaje = [types.Part.from_text(text=prompt_final)]
            if contexto_documentos_subidos:
                partes_mensaje.append(
                    types.Part.from_text(
                        text="\n\n".join(contexto_documentos_subidos)
                    )
                )
            partes_mensaje.extend(imagenes_subidas_parts)

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
                            response = chat.send_message(partes_mensaje)
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
