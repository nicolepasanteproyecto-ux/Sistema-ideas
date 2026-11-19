# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 11:55:46 2026

@author: practicante-im
"""
from pathlib import Path
from datetime import datetime
import shutil
import os

import pandas as pd
import streamlit as st
import plotly.express as px
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

st.set_page_config(
    page_title="Gestión de Ideas | Grupo Purdy",
    page_icon="💡",
    layout="wide",
    initial_sidebar_state="expanded",
)

CARPETA = Path(__file__).resolve().parent
ARCHIVO = CARPETA / "Base_Datos_Ideas.xlsx"
CARPETA_RESPALDOS = CARPETA / "Respaldos"

# Se conserva la contraseña que ya utilizabas.
# Más adelante conviene trasladarla a st.secrets.
CONTRASENA_ADMIN = "Purdy2026"

COLUMNAS = [
    "ID", "Fecha y hora", "Metodologia", "ID sesion", "Categoria",
    "Idea detallada", "Resultado fisico", "Estado digitalizacion",
    "Enlace", "Fecha modificacion", "Estado registro"
]

METODOLOGIAS = ["Brainwriting", "SCAMPER", "6 Sombreros", "Como lo haria"]
ESTADOS = ["Pendiente", "En proceso", "Completado"]
CATEGORIAS = [
    "Mejora de procesos", "Innovacion", "Calidad",
    "Productividad", "Seguridad", "Otra"
]

# =========================================================
# ESTILOS VISUALES
# =========================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background: #F4F7FA;
    color: #172B4D;
}

.block-container {
    padding-top: 1.4rem;
    padding-bottom: 2.5rem;
    max-width: 1450px;
}

/* Barra superior de Streamlit */
[data-testid="stHeader"] {
    background: transparent;
}

/* Barra lateral */
section[data-testid="stSidebar"] {
    background: #102A43;
}

section[data-testid="stSidebar"] > div {
    background: #102A43;
}

section[data-testid="stSidebar"] * {
    color: #FFFFFF;
}

section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
    color: #D7E3ED;
}

section[data-testid="stSidebar"] .stButton > button {
    width: 100%;
    min-height: 2.75rem;
    text-align: left;
    justify-content: flex-start;
    border-radius: 9px;
    border: 1px solid rgba(255,255,255,.12);
    background: #193B56;
    color: #FFFFFF !important;
    font-weight: 600;
    transition: all .2s ease;
}

section[data-testid="stSidebar"] .stButton > button:hover {
    background: #24516F;
    border-color: #5BC6A5;
    color: #FFFFFF !important;
}

section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #168567;
    border-color: #168567;
    color: #FFFFFF !important;
}

/* Campos dentro de la barra lateral */
section[data-testid="stSidebar"] input {
    color: #172B4D !important;
    background: #FFFFFF !important;
}

section[data-testid="stSidebar"] label {
    color: #FFFFFF !important;
}

/* Encabezado principal */
.hero {
    background: linear-gradient(115deg, #102A43 0%, #176B57 62%, #239879 100%);
    padding: 27px 32px;
    border-radius: 18px;
    color: white;
    margin-bottom: 23px;
    box-shadow: 0 9px 26px rgba(16,42,67,.12);
}

.hero h1 {
    color: #FFFFFF;
    font-size: 2rem;
    margin: 0 0 7px 0;
    font-weight: 800;
}

.hero p {
    color: #E0F4EC;
    margin: 0;
    font-size: 1rem;
}

.eyebrow {
    text-transform: uppercase;
    letter-spacing: 2px;
    font-size: .70rem;
    font-weight: 800;
    color: #B7E4D3;
    margin-bottom: 9px;
}

.section-note {
    color: #627D98;
    margin-top: -7px;
    margin-bottom: 20px;
}

/* Botones */
.stButton > button,
.stFormSubmitButton > button {
    border-radius: 10px;
    font-weight: 700;
    min-height: 2.65rem;
}

.stButton > button[kind="primary"],
.stFormSubmitButton > button[kind="primary"] {
    background: #168567;
    border-color: #168567;
    color: #FFFFFF !important;
}

/* Tarjetas de indicadores */
.metric-card {
    background: #FFFFFF;
    border: 1px solid #E5EBF0;
    padding: 18px 19px;
    border-radius: 14px;
    box-shadow: 0 3px 12px rgba(16,42,67,.04);
    min-height: 112px;
}

.metric-label {
    color: #627D98;
    font-weight: 600;
    font-size: .86rem;
    margin: 0;
}

.metric-value {
    color: #102A43;
    font-weight: 800;
    font-size: 1.9rem;
    margin: 4px 0 0 0;
}

/* Tablas y separadores */
div[data-testid="stDataFrame"] {
    border: 1px solid #E5EBF0;
    border-radius: 12px;
    overflow: hidden;
}

hr {
    border-color: #E1E8EF;
}

.small-muted {
    color: #829AB1;
    font-size: .82rem;
}

/* Formularios */
div[data-testid="stForm"] {
    background: #FFFFFF;
    border: 1px solid #E5EBF0;
    border-radius: 14px;
    padding: 20px;
}

/* Mensajes */
div[data-testid="stAlert"] {
    border-radius: 11px;
}

/* Ocultar el menú técnico de Streamlit, no los controles de navegación */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# FUNCIONES DE EXCEL
# Se conserva el funcionamiento local que ya tenía tu código.
# La conexión de Excel con la nube se trabajará por separado.
# =========================================================

def catalogos_df():
    return pd.DataFrame({
        "Metodologias": pd.Series(METODOLOGIAS),
        "Estados": pd.Series(ESTADOS),
        "Categorias": pd.Series(CATEGORIAS),
    })


def aplicar_estilo_excel(ruta_archivo):
    try:
        wb = openpyxl.load_workbook(ruta_archivo)

        font_header = Font(name="DM Sans", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill(
            start_color="102A43",
            end_color="102A43",
            fill_type="solid"
        )
        alignment_header = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        border_thin = Border(
            left=Side(style="thin", color="DEE2E6"),
            right=Side(style="thin", color="DEE2E6"),
            top=Side(style="thin", color="DEE2E6"),
            bottom=Side(style="thin", color="DEE2E6")
        )

        if "Ideas" in wb.sheetnames:
            ws = wb["Ideas"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(COLUMNAS))}{max(ws.max_row, 1)}"
            ws.row_dimensions[1].height = 32

            for col_num in range(1, len(COLUMNAS) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = alignment_header

            for row_idx in range(2, ws.max_row + 1):
                ws.row_dimensions[row_idx].height = 25
                for cell in ws[row_idx][:len(COLUMNAS)]:
                    cell.border = border_thin
                    cell.font = Font(name="DM Sans", size=10, color="212529")
                    cell.alignment = Alignment(vertical="center", wrap_text=True)

            anchos = {
                "A": 10, "B": 21, "C": 20, "D": 17, "E": 22,
                "F": 55, "G": 28, "H": 23, "I": 35, "J": 21, "K": 18
            }
            for columna, ancho in anchos.items():
                ws.column_dimensions[columna].width = ancho

            max_r = max(ws.max_row, 100)

            # Evita duplicar las validaciones cada vez que se aplica formato.
            ws.data_validations.dataValidation = []

            dv_metodo = DataValidation(
                type="list",
                formula1='"' + ",".join(METODOLOGIAS) + '"',
                allow_blank=True
            )
            ws.add_data_validation(dv_metodo)
            dv_metodo.add(f"C2:C{max_r}")

            dv_cat = DataValidation(
                type="list",
                formula1='"' + ",".join(CATEGORIAS) + '"',
                allow_blank=True
            )
            ws.add_data_validation(dv_cat)
            dv_cat.add(f"E2:E{max_r}")

            dv_est = DataValidation(
                type="list",
                formula1='"' + ",".join(ESTADOS) + '"',
                allow_blank=True
            )
            ws.add_data_validation(dv_est)
            dv_est.add(f"H2:H{max_r}")

        wb.save(ruta_archivo)

    except Exception as exc:
        print(f"No se pudo aplicar el formato al Excel: {exc}")


def crear_base_datos():
    if not ARCHIVO.exists():
        with pd.ExcelWriter(ARCHIVO, engine="openpyxl") as writer:
            pd.DataFrame(columns=COLUMNAS).to_excel(
                writer, sheet_name="Ideas", index=False
            )
            catalogos_df().to_excel(
                writer, sheet_name="Catalogos", index=False
            )
        aplicar_estilo_excel(ARCHIVO)


def cargar_datos():
    crear_base_datos()

    try:
        df = pd.read_excel(ARCHIVO, sheet_name="Ideas")

        for col in COLUMNAS:
            if col not in df.columns:
                df[col] = ""

        df = df[COLUMNAS].copy()
        df["Estado registro"] = df["Estado registro"].fillna("Activo")
        return df

    except PermissionError:
        st.error(
            "No se pudo leer el Excel. Cierra el archivo y vuelve a actualizar."
        )
        return pd.DataFrame(columns=COLUMNAS)

    except Exception as exc:
        st.error(f"No se pudo leer la hoja 'Ideas': {exc}")
        return pd.DataFrame(columns=COLUMNAS)


def guardar_datos(df):
    CARPETA_RESPALDOS.mkdir(exist_ok=True)

    if ARCHIVO.exists():
        sello = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        shutil.copy2(
            ARCHIVO,
            CARPETA_RESPALDOS / f"Base_Datos_Ideas_{sello}.xlsx"
        )

    temporal = CARPETA / "Base_Datos_Ideas_temporal.xlsx"

    try:
        with pd.ExcelWriter(temporal, engine="openpyxl") as writer:
            df.reindex(columns=COLUMNAS).to_excel(
                writer, sheet_name="Ideas", index=False
            )
            catalogos_df().to_excel(
                writer, sheet_name="Catalogos", index=False
            )

        os.replace(temporal, ARCHIVO)
        aplicar_estilo_excel(ARCHIVO)

    except PermissionError as exc:
        if temporal.exists():
            temporal.unlink(missing_ok=True)

        raise PermissionError(
            "Excel está abierto o bloqueado. Cierra "
            "Base_Datos_Ideas.xlsx e intenta de nuevo."
        ) from exc


def siguiente_id(df):
    if df.empty:
        return 1

    ids = pd.to_numeric(df["ID"], errors="coerce")
    return int(ids.max()) + 1 if ids.notna().any() else 1


def guardar_nueva_idea(datos):
    df_actual = cargar_datos()

    fila = {
        "ID": siguiente_id(df_actual),
        "Fecha y hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        **datos,
        "Fecha modificacion": "",
        "Estado registro": "Activo",
    }

    guardar_datos(
        pd.concat([df_actual, pd.DataFrame([fila])], ignore_index=True)
    )
    return fila["ID"]


def actualizar_idea(identificador, datos):
    df_actual = cargar_datos()

    mascara = (
        pd.to_numeric(df_actual["ID"], errors="coerce")
        == int(identificador)
    )

    if not mascara.any():
        raise ValueError("No se encontró el registro seleccionado.")

    for campo, valor in datos.items():
        df_actual.loc[mascara, campo] = valor

    df_actual.loc[mascara, "Fecha modificacion"] = (
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    )

    guardar_datos(df_actual)


def eliminar_idea(identificador):
    df_actual = cargar_datos()

    mascara = (
        pd.to_numeric(df_actual["ID"], errors="coerce")
        == int(identificador)
    )

    if not mascara.any():
        raise ValueError("No se encontró el registro seleccionado.")

    df_actual.loc[mascara, "Estado registro"] = "Eliminado"
    df_actual.loc[mascara, "Fecha modificacion"] = (
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    )

    guardar_datos(df_actual)


# =========================================================
# ESTADO DE SESIÓN Y NAVEGACIÓN
# Usuario: Inicio y Registrar ideas.
# Admin: todas las secciones.
# =========================================================

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if "opcion" not in st.session_state:
    st.session_state.opcion = "Menú principal"

if "mensaje_login" not in st.session_state:
    st.session_state.mensaje_login = ""

OPCIONES_USUARIO = [
    "Menú principal",
    "Sistema de ideas",
]

OPCIONES_ADMIN = [
    "Menú principal",
    "Sistema de ideas",
    "Dashboard",
    "Consultar ideas",
    "Gestionar ideas",
    "Base de datos",
]

ICONOS = {
    "Menú principal": "🏠",
    "Sistema de ideas": "💡",
    "Dashboard": "📊",
    "Consultar ideas": "🔎",
    "Gestionar ideas": "⚙️",
    "Base de datos": "🗄️",
}

with st.sidebar:
    st.markdown("## 💡 Gestión de Ideas")
    st.caption("PLATAFORMA DE INNOVACIÓN")
    st.divider()

    st.markdown("**MENÚ PRINCIPAL**")

    opciones_visibles = (
        OPCIONES_ADMIN if st.session_state.autenticado
        else OPCIONES_USUARIO
    )

    # Navegación por botones: elimina el selectbox y su campo de búsqueda.
    for nombre in opciones_visibles:
        seleccionado = st.session_state.opcion == nombre

        if st.button(
            f"{ICONOS[nombre]}  {nombre}",
            key=f"navegacion_{nombre}",
            use_container_width=True,
            type="primary" if seleccionado else "secondary",
        ):
            st.session_state.opcion = nombre
            st.rerun()

    st.divider()

    if not st.session_state.autenticado:
        with st.expander("🔐 Acceso Creador / Admin", expanded=False):
            password_input = st.text_input(
                "Contraseña",
                type="password",
                key="password_admin"
            )

            if st.button(
                "Ingresar",
                key="btn_login",
                use_container_width=True,
                type="primary"
            ):
                if password_input == CONTRASENA_ADMIN:
                    st.session_state.autenticado = True
                    st.session_state.opcion = "Dashboard"
                    st.session_state.mensaje_login = ""
                    st.rerun()
                else:
                    st.session_state.mensaje_login = "Contraseña incorrecta."

            if st.session_state.mensaje_login:
                st.error(st.session_state.mensaje_login)

    else:
        st.markdown("**PANEL DE ADMINISTRACIÓN**")
        st.caption("Acceso completo habilitado.")
        st.caption(f"📁 {ARCHIVO.name}")

        if st.button(
            "↻ Actualizar datos",
            key="btn_actualizar",
            use_container_width=True
        ):
            st.rerun()

        if st.button(
            "🔒 Cerrar sesión",
            key="btn_logout",
            use_container_width=True
        ):
            st.session_state.autenticado = False
            st.session_state.opcion = "Menú principal"
            st.session_state.mensaje_login = ""
            st.rerun()


# Validación adicional de permisos por si se manipula el estado de navegación.
if (
    not st.session_state.autenticado
    and st.session_state.opcion not in OPCIONES_USUARIO
):
    st.session_state.opcion = "Menú principal"

opcion = st.session_state.opcion

# Cargar datos después de preparar la navegación.
df = cargar_datos()

if "Estado registro" in df.columns:
    activos = df[
        df["Estado registro"].fillna("Activo") != "Eliminado"
    ].copy()
else:
    activos = pd.DataFrame(columns=COLUMNAS)


# =========================================================
# COMPONENTE DE ENCABEZADO
# =========================================================

def encabezado(titulo, descripcion, etiqueta="INNOVACIÓN · MEJORA CONTINUA"):
    st.markdown(
        f"""
        <div class="hero">
            <div class="eyebrow">{etiqueta}</div>
            <h1>{titulo}</h1>
            <p>{descripcion}</p>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# 1. INICIO
# =========================================================

if opcion == "Menú principal":
    encabezado(
        "💡 Sistema de Gestión de Ideas",
        "Un espacio para compartir, organizar e impulsar ideas de mejora e innovación."
    )

    st.markdown("### Bienvenido(a)")
    st.markdown(
        """
        Esta plataforma permite registrar ideas generadas por los equipos,
        dar seguimiento a su digitalización y facilitar su gestión.
        """
    )

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-label">💡 REGISTRO DE IDEAS</div>
                <h3 style="color:#102A43;margin-bottom:5px;">
                    Comparte una propuesta
                </h3>
                <p style="color:#627D98;margin:0;">
                    Registra ideas, oportunidades de mejora y propuestas
                    de innovación.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <div class="metric-card">
                <div class="metric-label">🤝 MEJORA CONTINUA</div>
                <h3 style="color:#102A43;margin-bottom:5px;">
                    Construimos mejores soluciones
                </h3>
                <p style="color:#627D98;margin:0;">
                    Las ideas contribuyen a identificar oportunidades
                    y fortalecer nuestro trabajo en equipo.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

    if st.session_state.autenticado:
        st.divider()
        st.success(
            "Sesión de administrador activa. "
            "Utiliza el menú lateral para gestionar el sistema."
        )
    else:
        st.info(
            "Selecciona «Registrar ideas» en el menú lateral "
            "para compartir una propuesta."
        )


# =========================================================
# 2. REGISTRO DE IDEAS
# =========================================================

elif opcion == "Sistema de ideas":
    encabezado(
        "💡 Registrar una idea",
        "Completa la información para compartir una nueva propuesta.",
        "INNOVACIÓN · REGISTRO DE IDEAS"
    )

    st.markdown("### Información de la idea")
    st.markdown(
        '<p class="section-note">Los campos marcados con * son obligatorios.</p>',
        unsafe_allow_html=True
    )

    with st.form("formulario_nueva_idea", clear_on_submit=True):
        a, b = st.columns(2)

        with a:
            metodologia = st.selectbox(
                "Metodología *",
                METODOLOGIAS
            )
            sesion = st.text_input(
                "ID de sesión",
                placeholder="Ej. SES-2026-001"
            )
            categoria = st.selectbox(
                "Categoría *",
                CATEGORIAS
            )

        with b:
            estado = st.selectbox(
                "Estado de digitalización",
                ESTADOS
            )
            resultado = st.text_input(
                "Resultado físico / referencia",
                placeholder="Ej. Hoja física #3"
            )
            enlace = st.text_input(
                "Enlace de referencia",
                placeholder="https://..."
            )

        idea = st.text_area(
            "Descripción detallada de la idea *",
            height=160,
            placeholder=(
                "Describe la situación u oportunidad, "
                "la propuesta y el beneficio esperado..."
            )
        )

        enviado = st.form_submit_button(
            "💾 Registrar idea",
            type="primary",
            use_container_width=True
        )

        if enviado:
            if not idea.strip():
                st.error("La descripción de la idea es obligatoria.")
            else:
                try:
                    nuevo_id = guardar_nueva_idea({
                        "Metodologia": metodologia,
                        "ID sesion": sesion.strip(),
                        "Categoria": categoria,
                        "Idea detallada": idea.strip(),
                        "Resultado fisico": resultado.strip(),
                        "Estado digitalizacion": estado,
                        "Enlace": enlace.strip(),
                    })

                    st.success(
                        f"¡Idea #{nuevo_id} registrada correctamente!"
                    )
                    st.toast("Registro guardado con éxito", icon="✅")

                except PermissionError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"No se pudo guardar la idea: {exc}")


# =========================================================
# 3. DASHBOARD - SOLO ADMIN
# =========================================================

elif opcion == "Dashboard" and st.session_state.autenticado:
    encabezado(
        "📊 Dashboard de ideas",
        "Indicadores generales y seguimiento de las ideas registradas.",
        "PANEL DE ADMINISTRACIÓN"
    )

    total = len(activos)
    completadas = (
        activos["Estado digitalizacion"] == "Completado"
    ).sum()
    proceso = (
        activos["Estado digitalizacion"] == "En proceso"
    ).sum()
    pendientes = (
        activos["Estado digitalizacion"] == "Pendiente"
    ).sum()

    porcentaje_completado = (
        int(completadas / total * 100) if total > 0 else 0
    )

    c1, c2, c3, c4 = st.columns(4)

    metricas = [
        (c1, "Total de ideas", total, "#102A43"),
        (c2, "Completadas", int(completadas), "#168567"),
        (c3, "En proceso", int(proceso), "#3197B5"),
        (c4, "Pendientes", int(pendientes), "#E5A93D"),
    ]

    for columna, etiqueta, valor, color in metricas:
        with columna:
            st.markdown(
                f"""
                <div class="metric-card" style="border-left:5px solid {color};">
                    <p class="metric-label">{etiqueta}</p>
                    <h3 class="metric-value">{valor}</h3>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown("")
    st.markdown(
        f"**Progreso de digitalización: {porcentaje_completado}%**"
    )
    st.progress(porcentaje_completado / 100 if total > 0 else 0)

    st.divider()

    if not activos.empty:
        izquierda, derecha = st.columns(2)

        with izquierda:
            resumen = (
                activos.groupby("Metodologia")
                .size()
                .reset_index(name="Cantidad")
            )

            fig = px.bar(
                resumen,
                x="Metodologia",
                y="Cantidad",
                color="Metodologia",
                title="Ideas por metodología",
                color_discrete_sequence=[
                    "#168567", "#3197B5", "#E5A93D", "#8067B7"
                ],
            )

            fig.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                font_color="#334E68",
                showlegend=False,
                margin=dict(l=15, r=15, t=55, b=15),
            )

            st.plotly_chart(fig, use_container_width=True)

        with derecha:
            resumen_estado = (
                activos.groupby("Estado digitalizacion")
                .size()
                .reset_index(name="Cantidad")
            )

            fig2 = px.pie(
                resumen_estado,
                names="Estado digitalizacion",
                values="Cantidad",
                title="Distribución por estado",
                hole=.58,
                color_discrete_sequence=[
                    "#E5A93D", "#3197B5", "#168567"
                ],
            )

            fig2.update_layout(
                paper_bgcolor="white",
                font_color="#334E68",
                margin=dict(l=15, r=15, t=55, b=15),
            )

            st.plotly_chart(fig2, use_container_width=True)

        st.subheader("Últimas ideas registradas")
        recientes = activos.tail(8).iloc[::-1]

        st.dataframe(
            recientes[
                [
                    "ID", "Fecha y hora", "Metodologia", "Categoria",
                    "Idea detallada", "Estado digitalizacion"
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

    else:
        st.info("Todavía no hay ideas registradas.")


# =========================================================
# 4. CONSULTAR IDEAS - SOLO ADMIN
# =========================================================

elif opcion == "Consultar ideas" and st.session_state.autenticado:
    encabezado(
        "🔎 Consultar ideas",
        "Busca, filtra y revisa los registros activos del sistema.",
        "PANEL DE ADMINISTRACIÓN"
    )

    if activos.empty:
        st.info("No hay ideas disponibles para consultar.")

    else:
        f1, f2, f3 = st.columns(3)

        with f1:
            metodo = st.selectbox(
                "Metodología",
                ["Todas"] + METODOLOGIAS
            )

        with f2:
            cat = st.selectbox(
                "Categoría",
                ["Todas"] + CATEGORIAS
            )

        with f3:
            estado_f = st.selectbox(
                "Estado",
                ["Todos"] + ESTADOS
            )

        busqueda = st.text_input(
            "Buscar idea",
            placeholder="Escribe una palabra clave, sesión o referencia..."
        )

        filtrado = activos.copy()

        if metodo != "Todas":
            filtrado = filtrado[
                filtrado["Metodologia"] == metodo
            ]

        if cat != "Todas":
            filtrado = filtrado[
                filtrado["Categoria"] == cat
            ]

        if estado_f != "Todos":
            filtrado = filtrado[
                filtrado["Estado digitalizacion"] == estado_f
            ]

        if busqueda.strip():
            campos = [
                "Idea detallada", "ID sesion",
                "Resultado fisico", "Enlace"
            ]

            mascara = pd.Series(False, index=filtrado.index)

            for campo in campos:
                mascara |= (
                    filtrado[campo]
                    .fillna("")
                    .astype(str)
                    .str.contains(busqueda, case=False, na=False)
                )

            filtrado = filtrado[mascara]

        st.caption(f"{len(filtrado)} registro(s) encontrado(s)")

        st.dataframe(
            filtrado.drop(columns=["Estado registro"]),
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# 5. GESTIONAR IDEAS - SOLO ADMIN
# =========================================================

elif opcion == "Gestionar ideas" and st.session_state.autenticado:
    encabezado(
        "⚙️ Gestionar ideas",
        "Edita registros existentes o retira ideas del sistema.",
        "PANEL DE ADMINISTRACIÓN"
    )

    if activos.empty:
        st.info("No hay ideas activas para gestionar.")

    else:
        opciones = {
            f"#{int(row['ID'])} · {str(row['Idea detallada'])[:75]}":
                int(row["ID"])
            for _, row in activos.iterrows()
        }

        etiqueta = st.selectbox(
            "Selecciona la idea que deseas gestionar",
            list(opciones.keys())
        )

        id_sel = opciones[etiqueta]

        registro = activos[
            pd.to_numeric(activos["ID"], errors="coerce") == id_sel
        ].iloc[0]

        st.markdown(
            f"**Idea seleccionada:** #{id_sel}"
        )

        tab_editar, tab_eliminar = st.tabs(
            ["✏️ Editar idea", "🗑️ Eliminar idea"]
        )

        with tab_editar:
            with st.form(f"editar_{id_sel}"):
                x1, x2 = st.columns(2)

                with x1:
                    m_idx = (
                        METODOLOGIAS.index(registro["Metodologia"])
                        if registro["Metodologia"] in METODOLOGIAS else 0
                    )

                    m_edit = st.selectbox(
                        "Metodología",
                        METODOLOGIAS,
                        index=m_idx
                    )

                    ses_edit = st.text_input(
                        "ID de sesión",
                        value=str(registro["ID sesion"] or "")
                    )

                    c_idx = (
                        CATEGORIAS.index(registro["Categoria"])
                        if registro["Categoria"] in CATEGORIAS else 0
                    )

                    c_edit = st.selectbox(
                        "Categoría",
                        CATEGORIAS,
                        index=c_idx
                    )

                with x2:
                    e_idx = (
                        ESTADOS.index(registro["Estado digitalizacion"])
                        if registro["Estado digitalizacion"] in ESTADOS else 0
                    )

                    e_edit = st.selectbox(
                        "Estado de digitalización",
                        ESTADOS,
                        index=e_idx
                    )

                    r_edit = st.text_input(
                        "Resultado físico / referencia",
                        value=str(registro["Resultado fisico"] or "")
                    )

                    l_edit = st.text_input(
                        "Enlace",
                        value=str(registro["Enlace"] or "")
                    )

                idea_edit = st.text_area(
                    "Descripción detallada",
                    value=str(registro["Idea detallada"] or ""),
                    height=150
                )

                actualizar = st.form_submit_button(
                    "💾 Guardar cambios",
                    type="primary",
                    use_container_width=True
                )

                if actualizar:
                    if not idea_edit.strip():
                        st.error(
                            "La descripción no puede quedar vacía."
                        )
                    else:
                        try:
                            actualizar_idea(
                                id_sel,
                                {
                                    "Metodologia": m_edit,
                                    "ID sesion": ses_edit.strip(),
                                    "Categoria": c_edit,
                                    "Idea detallada": idea_edit.strip(),
                                    "Resultado fisico": r_edit.strip(),
                                    "Estado digitalizacion": e_edit,
                                    "Enlace": l_edit.strip(),
                                }
                            )

                            st.success(
                                f"La idea #{id_sel} fue actualizada."
                            )
                            st.rerun()

                        except Exception as exc:
                            st.error(
                                f"No se pudo actualizar: {exc}"
                            )

        with tab_eliminar:
            st.warning(
                "La eliminación será lógica: el registro se conserva "
                "en Excel, pero deja de aparecer en consultas y dashboard."
            )

            confirmar = st.checkbox(
                f"Confirmo que deseo eliminar la idea #{id_sel}"
            )

            if st.button(
                "Eliminar idea seleccionada",
                type="primary",
                disabled=not confirmar
            ):
                try:
                    eliminar_idea(id_sel)

                    st.success(
                        f"La idea #{id_sel} fue marcada como eliminada."
                    )
                    st.rerun()

                except Exception as exc:
                    st.error(
                        f"No se pudo eliminar: {exc}"
                    )


# =========================================================
# 6. BASE DE DATOS - SOLO ADMIN
# =========================================================

elif opcion == "Base de datos" and st.session_state.autenticado:
    encabezado(
        "🗄️ Base de datos",
        "Consulta el contenido del archivo Excel y administra sus respaldos.",
        "PANEL DE ADMINISTRACIÓN"
    )

    st.caption(f"Ubicación del archivo: {ARCHIVO}")

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    if ARCHIVO.exists():
        with open(ARCHIVO, "rb") as archivo:
            st.download_button(
                "📥 Descargar una copia del Excel",
                data=archivo.read(),
                file_name="Base_Datos_Ideas.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    st.info(
        "Los registros y las ediciones se guardan en el archivo Excel "
        "local del entorno donde se ejecuta la aplicación."
    )

    st.markdown("### Respaldos")

    if CARPETA_RESPALDOS.exists():
        copias = sorted(
            CARPETA_RESPALDOS.glob("*.xlsx"),
            reverse=True
        )

        st.caption(
            f"{len(copias)} copia(s) disponible(s) en la carpeta Respaldos."
        )

        if copias:
            st.write(f"Último respaldo: {copias[0].name}")
    else:
        st.caption(
            "La carpeta Respaldos se creará al guardar el primer cambio."
        )


# =========================================================
# CONTROL DE ACCESO: SECCIÓN NO AUTORIZADA
# =========================================================

else:
    st.error(
        "No tienes permisos para acceder a esta sección. "
        "Inicia sesión como administrador."
    )
    st.session_state.opcion = "Menú principal"


# =========================================================
# PIE DE PÁGINA
# =========================================================

st.divider()

pie1, pie2 = st.columns([3, 1])

with pie1:
    st.caption("Sistema de Gestión de Ideas · Grupo Purdy")

with pie2:
    st.caption(datetime.now().strftime("%d/%m/%Y · %H:%M"))
