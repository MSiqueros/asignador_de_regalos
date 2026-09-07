"""Lectura y validación de los archivos Excel de entrada.

Se mantiene libre de Streamlit para poder probarse de forma aislada: las
funciones devuelven los mensajes de error y es la capa de UI la que decide
cómo mostrarlos.
"""
import pandas as pd

# Mapeo de las columnas del Excel de origen a los nombres internos.
INV_MAPPING = {
    "FECHACONTABILIZACION": "FechaIngreso",
    "ZONA": "ZonaElegible",
    "TIPOREGALO": "TipoRegalo",
    "ID": "CodigoArticulo",
    "OBSERVACION": "DescripcionArticulo",
    "CANTIDAD": "CantidadDisponible",
}
TDAS_MAPPING = {
    "CODIGO": "IDTienda",
    "NOMBRE_COLABORADOR": "NombreTienda",
    "TERRITORIO": "Zona",
    "TIPOREGALO": "TipoRegalo",
}

# Fila de encabezado de cada archivo (0-indexada).
FILA_ENCABEZADO_INVENTARIO = 2
FILA_ENCABEZADO_TIENDAS = 0


def preparar_dataframe(df, mapping, nombre_archivo):
    """Valida duplicados, renombra columnas y detecta las que falten.

    Devuelve (df_preparado, errores). Si `errores` no está vacío, el primer
    elemento es el mensaje principal y `df_preparado` es None.
    """
    duplicadas = df.columns[df.columns.duplicated()].tolist()
    if duplicadas:
        return None, [
            f"El archivo '{nombre_archivo}' tiene columnas duplicadas: {duplicadas}. "
            "Por favor, renómbralas o elimínalas en el archivo Excel original."
        ]

    # Sólo se renombran las columnas presentes cuyo destino aún no exista, para
    # no generar columnas duplicadas si el archivo ya trae el nombre interno.
    mapping_aplicable = {
        origen: destino
        for origen, destino in mapping.items()
        if origen in df.columns and destino not in df.columns
    }
    df = df.rename(columns=mapping_aplicable)

    faltantes = [col for col in mapping.values() if col not in df.columns]
    if faltantes:
        return None, [
            f"Al archivo '{nombre_archivo}' le faltan las siguientes columnas "
            f"esperadas: {faltantes}",
            f"Columnas encontradas: {list(df.columns)}",
        ]

    return df, []


def cargar_inventario(archivo):
    """Lee el Excel de inventario y devuelve (df, errores)."""
    df = pd.read_excel(archivo, header=FILA_ENCABEZADO_INVENTARIO)
    return preparar_dataframe(df, INV_MAPPING, "Inventario")


def cargar_tiendas(archivo):
    """Lee el Excel de tiendas y devuelve (df, errores)."""
    df = pd.read_excel(archivo, header=FILA_ENCABEZADO_TIENDAS)
    return preparar_dataframe(df, TDAS_MAPPING, "Tiendas")
