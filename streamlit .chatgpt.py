# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 11:55:46 2026

@author: practicante-im
"""
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