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
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Gestión de Ideas",
    page_icon="💡",
    layout="wide",
    initial_sidebar_state="expanded",
)

CARPETA = Path(__file__).resolve().parent
ARCHIVO = CARPETA / "Base_Datos_Ideas.xlsx"
CARPETA_RESPALDOS = CARPETA / "Respaldos"

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
# ESTILO VISUAL (STREAMLIT)
# =========================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #F5F7FA; color: #172B4D; }
[data-testid="stHeader"] { background: rgba(245,247,250,.92); }
.block-container { padding-top: 1.7rem; padding-bottom: 2.5rem; max-width: 1450px; }
section[data-testid="stSidebar"] { background: #102A43; }
section[data-testid="stSidebar"] * { color: #F3F7FA !important; }
section[data-testid="stSidebar"] [data-testid="stRadio"] label {
    background: rgba(255,255,255,.07); border-radius: 10px;
    padding: 9px 12px; margin: 4px 0;
}
.hero {
    background: linear-gradient(115deg, #102A43 0%, #176B57 58%, #239879 100%);
    padding: 30px 34px; border-radius: 18px; color: white;
    margin-bottom: 24px; box-shadow: 0 10px 28px rgba(16,42,67,.14);
}
.hero h1 { color: white; font-size: 2rem; margin: 0 0 7px 0; font-weight: 800; }
.hero p { color: #D9F3E9; margin: 0; font-size: 1rem; }
.eyebrow { text-transform: uppercase; letter-spacing: 2px; font-size: .72rem;
    font-weight: 700; color: #B7E4D3; margin-bottom: 9px; }
.section-note { color: #627D98; margin-top: -8px; margin-bottom: 18px; }
.stButton > button, .stFormSubmitButton > button {
    border-radius: 10px; font-weight: 700; min-height: 2.7rem;
}
.stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
    background: #168567; border-color: #168567;
}
div[data-testid="stDataFrame"] { border: 1px solid #E5EBF0; border-radius: 12px; overflow: hidden; }
hr { border-color: #E1E8EF; }
.small-muted { color: #829AB1; font-size: .82rem; }
</style>
""", unsafe_allow_html=True)


# =========================================================
# EXCEL: CREACIÓN, ESTILIZACIÓN EJECUTIVA Y ESCRITURA
# =========================================================

def catalogos_df():
    return pd.DataFrame({
        "Metodologias": pd.Series(METODOLOGIAS),
        "Estados": pd.Series(ESTADOS),
        "Categorias": pd.Series(CATEGORIAS),
    })

def aplicar_estilo_excel(ruta_archivo):
    """Aplica un diseño ejecutivo y profesional inspirado en bases de datos corporativas avanzadas."""
    try:
        wb = openpyxl.load_workbook(ruta_archivo)
        
        # Estilos ejecutivos (Gris plomo formal con texto blanco)
        font_header = Font(name="DM Sans", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="343A40", end_color="343A40", fill_type="solid")
        alignment_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        border_thin = Border(
            left=Side(style='thin', color='DEE2E6'),
            right=Side(style='thin', color='DEE2E6'),
            top=Side(style='thin', color='DEE2E6'),
            bottom=Side(style='thin', color='DEE2E6')
        )
        
        if "Ideas" in wb.sheetnames:
            ws = wb["Ideas"]
            ws.freeze_panes = "A2"
            
            # Altura y diseño de la cabecera
            ws.row_dimensions[1].height = 30
            for col_num in range(1, len(COLUMNAS) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = alignment_header

            # Estilo de celdas de datos con bordes finos y tipografía clara
            row_idx = 2
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(COLUMNAS)):
                ws.row_dimensions[row_idx].height = 22
                for cell in row:
                    cell.border = border_thin
                    cell.font = Font(name="DM Sans", size=10, color="212529")
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                row_idx += 1

            # Ancho de columnas automatizado y holgado para legibilidad
            for col in ws.columns:
                max_len = 0
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 5, 18)

            # Validaciones de datos (Listas desplegables en Excel)
            max_r = max(ws.max_row, 100)
            
            dv_metodo = DataValidation(type="list", formula1=f'"' + ",".join(METODOLOGIAS) + '"', allow_blank=True)
            ws.add_data_validation(dv_metodo)
            dv_metodo.add(f"C2:C{max_r}")

            dv_cat = DataValidation(type="list", formula1=f'"' + ",".join(CATEGORIAS) + '"', allow_blank=True)
            ws.add_data_validation(dv_cat)
            dv_cat.add(f"E2:E{max_r}")

            dv_est = DataValidation(type="list", formula1=f'"' + ",".join(ESTADOS) + '"', allow_blank=True)
            ws.add_data_validation(dv_est)
            dv_est.add(f"H2:H{max_r}")

        wb.save(ruta_archivo)
    except Exception as e:
        print(f"No se pudo aplicar el formato ejecutivo al Excel: {e}")

def crear_base_datos():
    if not ARCHIVO.exists():
        with pd.ExcelWriter(ARCHIVO, engine="openpyxl") as writer:
            pd.DataFrame(columns=COLUMNAS).to_excel(writer, sheet_name="Ideas", index=False)
            catalogos_df().to_excel(writer, sheet_name="Catalogos", index=False)
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
        st.error("No se pudo leer el Excel. Cierra el archivo y vuelve a actualizar.")
        return pd.DataFrame(columns=COLUMNAS)
    except Exception as exc:
        st.error(f"No se pudo leer la hoja 'Ideas': {exc}")
        return pd.DataFrame(columns=COLUMNAS)

def guardar_datos(df):
    CARPETA_RESPALDOS.mkdir(exist_ok=True)
    if ARCHIVO.exists():
        sello = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(ARCHIVO, CARPETA_RESPALDOS / f"Base_Datos_Ideas_{sello}.xlsx")

    temporal = CARPETA / "Base_Datos_Ideas_temporal.xlsx"
    try:
        with pd.ExcelWriter(temporal, engine="openpyxl") as writer:
            df.reindex(columns=COLUMNAS).to_excel(writer, sheet_name="Ideas", index=False)
            catalogos_df().to_excel(writer, sheet_name="Catalogos", index=False)
        os.replace(temporal, ARCHIVO)
        aplicar_estilo_excel(ARCHIVO)
    except PermissionError as exc:
        if temporal.exists():
            temporal.unlink(missing_ok=True)
        raise PermissionError(
            "Excel está abierto o bloqueado. Cierra Base_Datos_Ideas.xlsx e intenta de nuevo."
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
    guardar_datos(pd.concat([df_actual, pd.DataFrame([fila])], ignore_index=True))
    return fila["ID"]

def actualizar_idea(identificador, datos):
    df_actual = cargar_datos()
    mascara = pd.to_numeric(df_actual["ID"], errors="coerce") == int(identificador)
    if not mascara.any():
        raise ValueError("No se encontró el registro seleccionado.")
    for campo, valor in datos.items():
        df_actual.loc[mascara, campo] = valor
    df_actual.loc[mascara, "Fecha modificacion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    guardar_datos(df_actual)

def eliminar_idea(identificador):
    df_actual = cargar_datos()
    mascara = pd.to_numeric(df_actual["ID"], errors="coerce") == int(identificador)
    if not mascara.any():
        raise ValueError("No se encontró el registro seleccionado.")
    df_actual.loc[mascara, "Estado registro"] = "Eliminado"
    df_actual.loc[mascara, "Fecha modificacion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    guardar_datos(df_actual)


# =========================================================
# ENCABEZADO Y NAVEGACIÓN
# =========================================================

st.markdown("""
<div class="hero">
  <div class="eyebrow">Innovación · Mejora continua</div>
  <h1>💡 Sistema de Gestión de Ideas</h1>
  <p>Un espacio para capturar, organizar y convertir ideas en mejoras.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 💡 Gestión de Ideas")
    st.caption("Panel de trabajo local")
    opcion = st.radio(
        "MENÚ PRINCIPAL",
        ["Dashboard", "Registrar idea", "Consultar ideas", "Gestionar ideas", "Base de datos"],
        label_visibility="visible",
    )
    st.divider()
    st.markdown("**Almacenamiento**")
    st.caption("📗 Excel ejecutivo estilizado")
    st.caption(f"📁 {ARCHIVO.name}")
    st.caption("🔒 Los datos permanecen en esta computadora.")
    if st.button("↻ Actualizar datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

df = cargar_datos()
activos = df[df["Estado registro"].fillna("Activo") != "Eliminado"].copy()


# =========================================================
# DASHBOARD
# =========================================================

if opcion == "Dashboard":
    st.subheader("Resumen general")
    st.markdown('<p class="section-note">Vista rápida del estado y avance de las ideas digitalizadas.</p>', unsafe_allow_html=True)
    
    total = len(activos)
    completadas = (activos["Estado digitalizacion"] == "Completado").sum()
    proceso = (activos["Estado digitalizacion"] == "En proceso").sum()
    pendientes = (activos["Estado digitalizacion"] == "Pendiente").sum()
    
    porcentaje_completado = int((completadas / total * 100) if total > 0 else 0)

    # Tarjetas de Métricas Visuales Avanzadas (KPIs con diseño personalizado)
    c1, c2, c3, c4 = st.columns(4)
    
    with c1:
        st.markdown(f"""
            <div style="background: white; border: 1px solid #E5EBF0; padding: 18px; border-radius: 14px; box-shadow: 0 3px 12px rgba(16,42,67,.04); border-left: 5px solid #102A43;">
                <p style="color: #627D98; font-weight: 600; font-size: 0.85rem; margin: 0;">Total de ideas</p>
                <h3 style="color: #102A43; font-weight: 800; font-size: 1.8rem; margin: 5px 0 0 0;">{total}</h3>
            </div>
        """, unsafe_allow_html=True)
        
    with c2:
        st.markdown(f"""
            <div style="background: white; border: 1px solid #E5EBF0; padding: 18px; border-radius: 14px; box-shadow: 0 3px 12px rgba(16,42,67,.04); border-left: 5px solid #168567;">
                <p style="color: #627D98; font-weight: 600; font-size: 0.85rem; margin: 0;">Completadas</p>
                <h3 style="color: #168567; font-weight: 800; font-size: 1.8rem; margin: 5px 0 0 0;">{int(completadas)}</h3>
            </div>
        """, unsafe_allow_html=True)
        
    with c3:
        st.markdown(f"""
            <div style="background: white; border: 1px solid #E5EBF0; padding: 18px; border-radius: 14px; box-shadow: 0 3px 12px rgba(16,42,67,.04); border-left: 5px solid #3197B5;">
                <p style="color: #627D98; font-weight: 600; font-size: 0.85rem; margin: 0;">En proceso</p>
                <h3 style="color: #3197B5; font-weight: 800; font-size: 1.8rem; margin: 5px 0 0 0;">{int(proceso)}</h3>
            </div>
        """, unsafe_allow_html=True)
        
    with c4:
        st.markdown(f"""
            <div style="background: white; border: 1px solid #E5EBF0; padding: 18px; border-radius: 14px; box-shadow: 0 3px 12px rgba(16,42,67,.04); border-left: 5px solid #E5A93D;">
                <p style="color: #627D98; font-weight: 600; font-size: 0.85rem; margin: 0;">Pendientes</p>
                <h3 style="color: #E5A93D; font-weight: 800; font-size: 1.8rem; margin: 5px 0 0 0;">{int(pendientes)}</h3>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"**Progreso de digitalización y cierre ({porcentaje_completado}% completado)**")
    st.progress(porcentaje_completado / 100 if total > 0 else 0)
    st.divider()

    if not activos.empty:
        izquierda, derecha = st.columns(2)
        with izquierda:
            resumen = activos.groupby("Metodologia").size().reset_index(name="Cantidad")
            fig = px.bar(
                resumen, x="Metodologia", y="Cantidad", color="Metodologia",
                title="Ideas por metodología",
                color_discrete_sequence=["#168567", "#3197B5", "#E5A93D", "#8067B7"],
            )
            fig.update_layout(plot_bgcolor="white", paper_bgcolor="white", font_color="#334E68", showlegend=False, margin=dict(l=15, r=15, t=55, b=15))
            st.plotly_chart(fig, use_container_width=True)
        with derecha:
            resumen_estado = activos.groupby("Estado digitalizacion").size().reset_index(name="Cantidad")
            fig2 = px.pie(
                resumen_estado, names="Estado digitalizacion", values="Cantidad",
                title="Distribución por estado", hole=.58,
                color_discrete_sequence=["#E5A93D", "#3197B5", "#168567"],
            )
            fig2.update_layout(paper_bgcolor="white", font_color="#334E68", margin=dict(l=15, r=15, t=55, b=15))
            st.plotly_chart(fig2, use_container_width=True)
        st.subheader("Últimas ideas registradas")
        recientes = activos.tail(8).iloc[::-1]
        st.dataframe(
            recientes[["ID", "Fecha y hora", "Metodologia", "Categoria", "Idea detallada", "Estado digitalizacion"]],
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Todavía no hay ideas registradas. Comienza desde «Registrar idea».")


# =========================================================
# REGISTRO
# =========================================================

elif opcion == "Registrar idea":
    st.subheader("Registrar una nueva idea")
    st.markdown('<p class="section-note">Completa los datos. Al guardar, el registro se formatea con estilo ejecutivo y se agrega al Excel local.</p>', unsafe_allow_html=True)
    with st.form("formulario_nueva_idea", clear_on_submit=True):
        a, b = st.columns(2)
        with a:
            metodologia = st.selectbox("Metodología *", METODOLOGIAS)
            sesion = st.text_input("ID de sesión", placeholder="Ej. SES-2026-001")
            categoria = st.selectbox("Categoría *", CATEGORIAS)
        with b:
            estado = st.selectbox("Estado de digitalización", ESTADOS)
            resultado = st.text_input("Resultado físico / referencia", placeholder="Ej. Hoja física #3")
            enlace = st.text_input("Enlace de referencia", placeholder="https://...")
        idea = st.text_area("Descripción detallada de la idea *", height=160, placeholder="Describe el problema, la propuesta y el beneficio esperado...")
        enviado = st.form_submit_button("💾 Guardar idea en Excel", type="primary", use_container_width=True)
        if enviado:
            if not idea.strip():
                st.error("La descripción de la idea es obligatoria.")
            else:
                try:
                    nuevo_id = guardar_nueva_idea({
                        "Metodologia": metodologia, "ID sesion": sesion.strip(),
                        "Categoria": categoria, "Idea detallada": idea.strip(),
                        "Resultado fisico": resultado.strip(), "Estado digitalizacion": estado,
                        "Enlace": enlace.strip(),
                    })
                    st.success(f"¡Idea #{nuevo_id} guardada y estilizada correctamente en {ARCHIVO.name}!")
                    st.toast("Registro guardado en Excel", icon="✅")
                except PermissionError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"No se pudo guardar la idea: {exc}")


# =========================================================
# CONSULTAS
# =========================================================

elif opcion == "Consultar ideas":
    st.subheader("Consultar ideas")
    st.markdown('<p class="section-note">Busca y filtra los registros activos guardados en Excel.</p>', unsafe_allow_html=True)
    if activos.empty:
        st.info("No hay ideas disponibles para consultar.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1: metodo = st.selectbox("Metodología", ["Todas"] + METODOLOGIAS)
        with f2: cat = st.selectbox("Categoría", ["Todas"] + CATEGORIAS)
        with f3: estado_f = st.selectbox("Estado", ["Todos"] + ESTADOS)
        busqueda = st.text_input("Buscar en la descripción, sesión o resultado", placeholder="Escribe una palabra clave...")
        filtrado = activos.copy()
        if metodo != "Todas": filtrado = filtrado[filtrado["Metodologia"] == metodo]
        if cat != "Todas": filtrado = filtrado[filtrado["Categoria"] == cat]
        if estado_f != "Todos": filtrado = filtrado[filtrado["Estado digitalizacion"] == estado_f]
        if busqueda.strip():
            campos = ["Idea detallada", "ID sesion", "Resultado fisico", "Enlace"]
            mascara = pd.Series(False, index=filtrado.index)
            for campo in campos:
                mascara |= filtrado[campo].fillna("").astype(str).str.contains(busqueda, case=False, na=False)
            filtrado = filtrado[mascara]
        st.caption(f"{len(filtrado)} registro(s) encontrado(s)")
        st.dataframe(filtrado.drop(columns=["Estado registro"]), use_container_width=True, hide_index=True)


# =========================================================
# EDICIÓN Y ELIMINACIÓN LÓGICA
# =========================================================

elif opcion == "Gestionar ideas":
    st.subheader("Gestionar ideas")
    st.markdown('<p class="section-note">Edita un registro existente o márcalo como eliminado manteniendo el diseño en Excel.</p>', unsafe_allow_html=True)
    if activos.empty:
        st.info("No hay ideas activas para gestionar.")
    else:
        opciones = {
            f"#{int(row['ID'])} · {str(row['Idea detallada'])[:75]}": int(row["ID"])
            for _, row in activos.iterrows()
        }
        etiqueta = st.selectbox("Selecciona la idea", list(opciones.keys()))
        id_sel = opciones[etiqueta]
        registro = activos[pd.to_numeric(activos["ID"], errors="coerce") == id_sel].iloc[0]
        tab_editar, tab_eliminar = st.tabs(["✏️ Editar idea", "🗑️ Eliminar idea"])
        with tab_editar:
            with st.form(f"editar_{id_sel}"):
                x1, x2 = st.columns(2)
                with x1:
                    m_idx = METODOLOGIAS.index(registro["Metodologia"]) if registro["Metodologia"] in METODOLOGIAS else 0
                    m_edit = st.selectbox("Metodología", METODOLOGIAS, index=m_idx)
                    ses_edit = st.text_input("ID de sesión", value=str(registro["ID sesion"] or ""))
                    c_idx = CATEGORIAS.index(registro["Categoria"]) if registro["Categoria"] in CATEGORIAS else 0
                    c_edit = st.selectbox("Categoría", CATEGORIAS, index=c_idx)
                with x2:
                    e_idx = ESTADOS.index(registro["Estado digitalizacion"]) if registro["Estado digitalizacion"] in ESTADOS else 0
                    e_edit = st.selectbox("Estado de digitalización", ESTADOS, index=e_idx)
                    r_edit = st.text_input("Resultado físico / referencia", value=str(registro["Resultado fisico"] or ""))
                    l_edit = st.text_input("Enlace", value=str(registro["Enlace"] or ""))
                idea_edit = st.text_area("Descripción detallada", value=str(registro["Idea detallada"] or ""), height=150)
                actualizar = st.form_submit_button("Guardar cambios en Excel", type="primary", use_container_width=True)
                if actualizar:
                    if not idea_edit.strip():
                        st.error("La descripción no puede quedar vacía.")
                    else:
                        try:
                            actualizar_idea(id_sel, {
                                "Metodologia": m_edit, "ID sesion": ses_edit.strip(),
                                "Categoria": c_edit, "Idea detallada": idea_edit.strip(),
                                "Resultado fisico": r_edit.strip(), "Estado digitalizacion": e_edit, "Enlace": l_edit.strip(),
                            })
                            st.success(f"La idea #{id_sel} fue actualizada.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"No se pudo actualizar: {exc}")
        with tab_eliminar:
            st.warning("La eliminación será lógica: el registro se conserva en Excel, pero deja de aparecer en las consultas y el dashboard.")
            confirmar = st.checkbox(f"Confirmo que deseo eliminar la idea #{id_sel}")
            if st.button("Eliminar idea seleccionada", type="primary", disabled=not confirmar):
                try:
                    eliminar_idea(id_sel)
                    st.success(f"La idea #{id_sel} fue marcada como eliminada.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"No se pudo eliminar: {exc}")


# =========================================================
# BASE DE DATOS Y DESCARGA
# =========================================================

elif opcion == "Base de datos":
    st.subheader("Base de datos local ejecutiva")
    st.markdown('<p class="section-note">Vista completa del contenido del archivo Excel con diseño corporativo formal.</p>', unsafe_allow_html=True)
    st.caption(f"Ubicación del archivo: {ARCHIVO}")
    st.dataframe(df, use_container_width=True, hide_index=True)
    if ARCHIVO.exists():
        with open(ARCHIVO, "rb") as archivo:
            st.download_button(
                "📥 Descargar una copia del Excel",
                data=archivo.read(),
                file_name="Base_Datos_Ideas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    st.info("Los nuevos registros y ediciones se guardan automáticamente aplicando diseño formal ejecutivo y listas desplegables.")
    st.markdown("**Respaldos automáticos**")
    if CARPETA_RESPALDOS.exists():
        copias = sorted(CARPETA_RESPALDOS.glob("*.xlsx"), reverse=True)
        st.caption(f"{len(copias)} copia(s) disponible(s) en la carpeta Respaldos.")
        if copias: st.write(copias[0].name)
    else:
        st.caption("Se creará la carpeta Respaldos al guardar el primer cambio.")


# =========================================================
# PIE
# =========================================================

st.divider()
pie1, pie2 = st.columns([3, 1])
with pie1: st.caption("Sistema de Gestión de Ideas · Excel ejecutivo local")
with pie2: st.caption(datetime.now().strftime("%d/%m/%Y · %H:%M"))