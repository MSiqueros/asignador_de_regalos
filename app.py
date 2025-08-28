# streamlit_app.py
import streamlit as st
import pandas as pd
from asignador_regalos import ejecutar_asignacion

st.set_page_config(layout="wide")
st.title("Asignación automatizada de regalos a tiendas")

# --- Columnas de la interfaz ---
col1, col2 = st.columns(2)

with col1:
    st.info("Carga los archivos de inventario y tiendas para comenzar.")
    inv_file = st.file_uploader("1. Inventario (inventario.xlsx)", type=["xlsx"])
    tdas_file = st.file_uploader("2. Tiendas (tiendas.xlsx)", type=["xlsx"])

with col2:
    st.info("Configura los parámetros de la asignación.")
    estrategia = st.selectbox(
        "Estrategia de asignación",
        ["Sobrantes", "Novedades", "AltoStock", "Equitativo"],
        help="Define qué artículos se usarán primero."
    )
    n_regalos = st.selectbox(
        "N° de regalos por tienda",
        [1, 2],
        help="Cuántos regalos se asignarán a cada tienda."
    )

# --- Función de preparación de datos ---
def preparar_dataframe(df, mapping, nombre_archivo):
    """
    Valida duplicados, renombra columnas y muestra un error si faltan columnas.
    """
    # 1. Validar columnas duplicadas en el archivo original
    columnas_duplicadas = df.columns[df.columns.duplicated()].tolist()
    if columnas_duplicadas:
        st.error(
            f"El archivo '{nombre_archivo}' tiene columnas duplicadas: {columnas_duplicadas}. "
            "Por favor, renómbralas o elimínalas en el archivo Excel original."
        )
        return None

    # 2. Renombrar columnas según el mapeo
    df = df.rename(columns=mapping)

    # 3. Chequear que todas las columnas requeridas existan después de renombrar
    columnas_requeridas = list(mapping.values())
    faltantes = [col for col in columnas_requeridas if col not in df.columns]
    if faltantes:
        st.error(f"Al archivo '{nombre_archivo}' le faltan las siguientes columnas esperadas: {faltantes}")
        return None

    return df

# --- Lógica principal de la aplicación ---
if st.button("🚀 Generar Asignación", use_container_width=True):
    if not inv_file or not tdas_file:
        st.error("⚠️ Sube los dos archivos requeridos antes de continuar.")
    else:
        with st.spinner("Procesando archivos y realizando asignaciones..."):
            try:
                # Leer archivos
                inv_raw = pd.read_excel(inv_file, header=2)
                tdas_raw = pd.read_excel(tdas_file, header=0)

                # --- MAPEADO DE COLUMNAS ACTUALIZADO ---
                inv_mapping = {
                "FECHACONTABILIZACION": "FechaIngreso",
                "ZONA": "ZonaElegible",
                "TIPOREGALO": "TipoRegalo",
                "ID": "CodigoArticulo",           # CORRECCIÓN: 'id' cambiado a 'ID'
                "OBSERVACION": "DescripcionArticulo", # CORRECCIÓN: 'observacion' cambiado a 'OBSERVACION'
                "CANTIDAD": "CantidadDisponible"
                }
                tdas_mapping = {
                    "CODIGO": "IDTienda",
                    "NOMBRE_COLABORADOR": "NombreTienda",
                    "TERRITORIO": "Zona",
                    "TIPOREGALO": "TipoRegalo"
                }

                # Preparar y validar DataFrames
                inv = preparar_dataframe(inv_raw, inv_mapping, "Inventario")
                tdas = preparar_dataframe(tdas_raw, tdas_mapping, "Tiendas")

                if inv is None or tdas is None:
                    st.stop()

                # Ejecutar la lógica de asignación
                asignaciones, inv_rest, reporte_txt, excel_bytes = ejecutar_asignacion(
                    inv, tdas, n_regalos, estrategia
                )

                st.success("¡Asignación completada con éxito!")

                # --- Pestañas de resultados ---
                tab1, tab2, tab3, tab4 = st.tabs(["🎯 Asignaciones", "📦 Inventario Restante", "📋 Reporte", "📊 Vistas Previas"])

                with tab1:
                    st.dataframe(asignaciones)
                    st.download_button(
                        "⬇️ Descargar asignacion_final.xlsx",
                        data=excel_bytes,
                        file_name="asignacion_final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                with tab2:
                    st.dataframe(inv_rest)

                with tab3:
                    st.text_area("Resumen de la ejecución", reporte_txt, height=300)
                    st.download_button(
                        "⬇️ Descargar reporte.txt",
                        data=reporte_txt.encode("utf-8"),
                        file_name="reporte.txt",
                        mime="text/plain",
                        use_container_width=True
                    )

                with tab4:
                    st.subheader("Vista previa del Inventario (datos normalizados)")
                    st.dataframe(inv.head())
                    st.subheader("Vista previa de Tiendas (datos normalizados)")
                    st.dataframe(tdas.head())

            except Exception as e:
                st.error(f"Ocurrió un error inesperado durante el proceso: {e}")
