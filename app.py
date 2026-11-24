import os
import streamlit as st
from crewai import Agent, Task, Crew, Process, LLM
from crewai.tools import tool
from duckduckgo_search import DDGS

st.set_page_config(
    page_title="APLICACIÓN DE UNA AGENTE IA | ODS 4",
    page_icon="🎓",
    layout="wide"
)

@tool("Buscador Web Educativo")
def buscar_web_educativa(consulta: str) -> str:
    """Busca estándares curriculares, guías de la UNESCO y recursos del ODS 4 en internet."""
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(consulta, max_results=4))
        if not resultados:
            return "No se encontraron resultados externos; usar base pedagógica interna."
        return "\n".join(
            [f"- {r['title']}: {r['body']} (Fuente: {r['href']})" for r in resultados]
        )
    except Exception as e:
        return f"Búsqueda web limitada temporalmente ({str(e)}). Continuar con conocimiento curricular."

with st.sidebar:
    st.title("⚙️ Núcleo del Agente")
    st.caption("Proyecto Universitario: Aplicación de una Agente IA (ODS 4)")
    api_key = st.text_input(
        "Google AI Studio API Key (Gratis)",
        type="password",
        help="Pega aquí tu clave generada en Google AI Studio"
    )
    st.divider()
    st.markdown("**Agentes Activos en el Pipeline:**")
    st.markdown("1. 🔍 **Investigador Curricular** (Web Search)")
    st.markdown("2. 🧠 **Arquitecto Pedagógico** (Adaptación ODS 4)")
    st.markdown("3. 🛡️ **Auditor de Equidad** (Control de Calidad)")

st.title("🎓 APLICACIÓN DE UNA AGENTE IA — ODS 4: Educación de Calidad")
st.markdown(
    "Sistema multi-agente autónomo enfocado en las **Metas 4.1 y 4.5 de la UNESCO**: "
    "aprendizaje efectivo, equitativo y adaptado a contextos con brechas educativas."
)

col1, col2, col3 = st.columns(3)
with col1:
    tema = st.text_input("Tema o competencia académica", value="Comprensión lectora y pensamiento crítico")
with col2:
    nivel = st.selectbox("Nivel educativo", ["Primaria (5.º - 6.º)", "Secundaria (1.º - 3.º)", "Secundaria (4.º - 5.º)", "Nivelación Preuniversitaria"])
with col3:
    contexto = st.selectbox(
        "Contexto de inclusión (Meta 4.5)",
        [
            "Zona rural con conectividad limitada (material imprimible)",
            "Estudiantes con rezago escolar y baja comprensión base",
            "Aula multigrado con recursos tecnológicos mínimos",
            "Estudiantes neurodivergentes (lectura fácil y estructura visual)"
        ]
    )

objetivo_extra = st.text_area(
    "Directriz general para el equipo de agentes",
    value="Diseña un módulo completo de nivelación de 90 minutos con lectura adaptada, analogías cotidianas, ejercicios progresivos y rúbrica de evaluación."
)

if st.button("🚀 Iniciar Ejecución Autónoma de Agentes", type="primary", use_container_width=True):
    if not api_key:
        st.error("Ingresa tu API Key gratuita de Google AI Studio en el panel izquierdo para activar los agentes.")
        st.stop()

    os.environ["GEMINI_API_KEY"] = api_key
    llm_gemini = LLM(
        model="gemini/gemini-2.5-flash",
        api_key=api_key,
        temperature=0.4
    )

    investigador = Agent(
        role="Investigador de Estándares Curriculares y ODS 4",
        goal="Investigar criterios pedagógicos oficiales y lineamientos de la UNESCO relevantes para el tema y contexto asignado.",
        backstory="Especialista en políticas educativas internacionales y diseño curricular basado en evidencia.",
        tools=[buscar_web_educativa],
        llm=llm_gemini,
        verbose=True,
        allow_delegation=False
    )

    arquitecto = Agent(
        role="Arquitecto Pedagógico Inclusivo",
        goal="Construir una guía didáctica completa, accesible y adaptada exactamente a las barreras del contexto del estudiante.",
        backstory="Experto en pedagogía diferenciada y diseño instruccional para poblaciones vulnerables (Meta 4.5 ODS 4).",
        llm=llm_gemini,
        verbose=True,
        allow_delegation=False
    )

    auditor = Agent(
        role="Auditor de Calidad Educativa y Equidad ODS 4",
        goal="Evaluar críticamente el módulo diseñado, corregir debilidades pedagógicas, añadir una rúbrica y emitir la versión final certificada.",
        backstory="Director académico riguroso que verifica que ningún material sea genérico y que cumpla indicadores reales de aprendizaje.",
        llm=llm_gemini,
        verbose=True,
        allow_delegation=False
    )

    tarea_1 = Task(
        description=(
            f"Investiga en internet estándares educativos, competencias clave y enfoques inclusivos sobre '{tema}' "
            f"para el nivel '{nivel}' considerando el contexto '{contexto}'. Resume los hallazgos clave y fuentes."
        ),
        expected_output="Informe técnico breve con competencias curriculares, criterios ODS 4 y fuentes consultadas.",
        agent=investigador
    )

    tarea_2 = Task(
        description=(
            f"Con base en la investigación previa y la directriz '{objetivo_extra}', redacta el módulo educativo completo "
            f"sobre '{tema}' para '{nivel}' en contexto '{contexto}'. Incluye: 1) Objetivos de aprendizaje, 2) Explicación con analogías cotidianas, "
            f"3) Actividad práctica sin depender de internet, 4) Tres ejercicios de dificultad progresiva con solucionario."
        ),
        expected_output="Borrador completo y estructurado del módulo educativo adaptado.",
        agent=arquitecto,
        context=[tarea_1]
    )

    tarea_3 = Task(
        description=(
            f"Audita el módulo educativo generado para '{contexto}'. Verifica si realmente elimina barreras de aprendizaje "
            f"según la Meta 4.5 del ODS 4. Mejora las actividades si son genéricas, añade una tabla de rúbrica de evaluación "
            f"(Inicio, En Proceso, Logrado) y encabeza el documento con un 'Índice de Alineación ODS 4' (puntaje sobre 100 y justificación)."
        ),
        expected_output="Módulo educativo final auditado, mejorado, con rúbrica de evaluación y reporte de cumplimiento ODS 4 en formato Markdown.",
        agent=auditor,
        context=[tarea_2]
    )

    with st.status("🤖 Agentes operando autónomamente... (Observa el flujo en tiempo real)", expanded=True) as status:
        st.write("⏳ **Fase 1:** El *Agente Investigador* está consultando estándares curriculares en la web...")

        def registrar_avance(output):
            st.write(f"✅ **Tarea completada por:** `{output.agent}` — Transfiriendo resultados al siguiente agente...")

        equipo_agentes = Crew(
            agents=[investigador, arquitecto, auditor],
            tasks=[tarea_1, tarea_2, tarea_3],
            process=Process.sequential,
            task_callback=registrar_avance
        )

        resultado_final = equipo_agentes.kickoff()
        status.update(label="✨ Ejecución multi-agente finalizada y auditada con éxito", state="complete", expanded=False)

    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Agentes Ejecutados", "3 Autónomos")
    m2.metric("Búsqueda Externa", "Activa (Web)")
    m3.metric("Meta ODS 4", "4.1 y 4.5 Validada")

    st.subheader("📄 Módulo Educativo Generado y Auditado")
    st.markdown(str(resultado_final))

    st.download_button(
        label="⬇️ Descargar Módulo Completo (.md)",
        data=str(resultado_final),
        file_name=f"Modulo_ODS4_{tema.replace(' ', '_')}.md",
        mime="text/markdown"
    )
