"""Lectura y validación de los archivos Excel de entrada.

Se mantiene libre de Streamlit para poder probarse de forma aislada: las
funciones devuelven los mensajes de error y es la capa de UI la que decide
cómo mostrarlos.
"""
import unicodedata

import pandas as pd

# Nombre interno -> nombres aceptados en el Excel de origen, en orden de
# preferencia. La comparación ignora mayúsculas, acentos, espacios y
# separadores, así que 'tamaño', 'TAMAÑO' y 'Tamano' se reconocen igual.
INV_COLUMNAS = {
    "FechaIngreso": ("FECHACONTABILIZACION", "FECHAINGRESO"),
    "ZonaElegible": ("ZONA", "ZONAELEGIBLE"),
    "TipoRegalo": ("TIPOREGALO",),
    "CodigoArticulo": ("ID",),
    "DescripcionArticulo": ("OBSERVACION", "DESCRIPCIONARTICULO"),
    "CantidadDisponible": ("CANTIDAD", "SALDO"),
}
TDAS_COLUMNAS = {
    "IDTienda": ("CODIGO",),
    "NombreTienda": ("NOMBRE_COLABORADOR",),
    "Zona": ("TERRITORIO", "ZONA"),
    # La plantilla de tiendas llama 'tamaño' al segmento que el inventario
    # llama 'TIPOREGALO'; es la columna con la que se cruzan ambos archivos.
    "TipoRegalo": ("TAMAÑO", "TIPOREGALO"),
}

# Columnas que enriquecen la asignación pero cuya ausencia no invalida el
# archivo: si no están, se crean vacías y el proceso continúa.
TDAS_COLUMNAS_OPCIONALES = {
    # Tipo de regalo del segundo obsequio. Vacío ⇒ la tienda recibe uno solo.
    "TipoRegaloAdicional": ("REGALO ADICIONAL", "TIPO REGALO ADICIONAL"),
}

# Fila de encabezado de cada archivo (0-indexada).
FILA_ENCABEZADO_INVENTARIO = 2
FILA_ENCABEZADO_TIENDAS = 0


def normalizar_encabezado(nombre):
    """Clave de comparación de encabezados: sin acentos, espacios ni separadores."""
    texto = unicodedata.normalize("NFKD", str(nombre))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return "".join(texto.split()).replace("_", "").replace("-", "").upper()


def _buscar_origen(alias, indice):
    """Nombre real del primer alias presente en `indice`, o None si no está.

    El índice se recibe como parámetro porque cambia entre pasadas: la de
    obligatorias mira los encabezados originales y la de opcionales los ya
    renombrados.
    """
    return next(
        (
            indice[normalizar_encabezado(a)]
            for a in alias
            if normalizar_encabezado(a) in indice
        ),
        None,
    )


def preparar_dataframe(df, columnas, nombre_archivo, opcionales=None):
    """Valida duplicados, renombra columnas y detecta las que falten.

    `opcionales` usa el mismo formato que `columnas`, pero su ausencia no es un
    error: la columna se crea vacía y su nombre queda en
    `df.attrs["columnas_opcionales_ausentes"]`.

    Devuelve (df_preparado, errores). Si `errores` no está vacío, el primer
    elemento es el mensaje principal y `df_preparado` es None.
    """
    duplicadas = df.columns[df.columns.duplicated()].tolist()
    if duplicadas:
        return None, [
            f"El archivo '{nombre_archivo}' tiene columnas duplicadas: {duplicadas}. "
            "Por favor, renómbralas o elimínalas en el archivo Excel original."
        ]

    # Encabezado normalizado -> nombre real. `setdefault` conserva la primera
    # aparición, que es la que gana si dos encabezados normalizan igual.
    presentes = {}
    for real in df.columns:
        presentes.setdefault(normalizar_encabezado(real), real)

    renombres, faltantes = {}, []
    for destino, alias in columnas.items():
        # Si el archivo ya trae el nombre interno, renombrar crearía un duplicado.
        if destino in df.columns:
            continue
        origen = _buscar_origen(alias, presentes)
        if origen is None:
            faltantes.append((destino, alias))
        else:
            renombres[origen] = destino

    if faltantes:
        errores = [
            f"Al archivo '{nombre_archivo}' le faltan las siguientes columnas "
            f"esperadas: {[destino for destino, _ in faltantes]}"
        ]
        for destino, alias in faltantes:
            errores.append(f"  · {destino}: se esperaba una de {list(alias)}")
        errores.append(f"Columnas encontradas: {list(df.columns)}")
        return None, errores

    df = df.rename(columns=renombres)

    # `presentes` mira los nombres previos al renombre; para las opcionales hay
    # que reindexar, o un alias que ya se llevó una obligatoria se perdería.
    disponibles = {}
    for real in df.columns:
        disponibles.setdefault(normalizar_encabezado(real), real)

    # Las opcionales se resuelven después de las obligatorias: si falta alguna
    # se crea vacía y se deja constancia, para que la UI pueda avisarlo.
    ausentes = []
    for destino, alias in (opcionales or {}).items():
        if destino not in df.columns:
            origen = _buscar_origen(alias, disponibles)
            if origen is None:
                df[destino] = ""
                ausentes.append(destino)
            else:
                df = df.rename(columns={origen: destino})
        # Unifica el vacío: una plantilla con la columna en blanco se lee como
        # NaN, y quien consuma la columna espera el mismo centinela siempre.
        df[destino] = df[destino].fillna("")

    df.attrs["columnas_opcionales_ausentes"] = ausentes
    return df, []


def elegir_hoja(libro, columnas, fila_encabezado):
    """Nombre de la hoja cuyo encabezado calza mejor con las columnas esperadas.

    Los archivos reales traen hojas auxiliares (tablas dinámicas, notas) que
    pueden quedar en primer lugar; `pd.read_excel` sin `sheet_name` tomaría
    esa y el mapeo fallaría con encabezados sin sentido.
    """
    aceptados = {normalizar_encabezado(a) for alias in columnas.values() for a in alias}
    aceptados |= {normalizar_encabezado(destino) for destino in columnas}

    mejor, mejor_puntaje = libro.sheet_names[0], -1
    for hoja in libro.sheet_names:
        try:
            encabezados = pd.read_excel(
                libro, sheet_name=hoja, header=fila_encabezado, nrows=0
            ).columns
        except (ValueError, IndexError):
            # Hoja con menos filas que `fila_encabezado`: no puede ser la buena.
            continue
        puntaje = sum(1 for c in encabezados if normalizar_encabezado(c) in aceptados)
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = hoja, puntaje
    return mejor


def _cargar(archivo, columnas, fila_encabezado, nombre_archivo, opcionales=None):
    """Abre el libro una sola vez, elige la hoja correcta y aplica el mapeo."""
    # Las opcionales también puntúan al elegir la hoja: son señal de cuál es
    # la buena, no ruido a ignorar.
    columnas_puntaje = {**columnas, **(opcionales or {})}
    with pd.ExcelFile(archivo) as libro:
        hojas = list(libro.sheet_names)
        hoja = elegir_hoja(libro, columnas_puntaje, fila_encabezado)
        df = pd.read_excel(libro, sheet_name=hoja, header=fila_encabezado)

    df, errores = preparar_dataframe(df, columnas, nombre_archivo, opcionales)
    if df is not None:
        # Informativo para la UI: deja constancia de qué hoja se leyó.
        df.attrs["hoja"] = hoja
    else:
        errores.append(f"Hoja leída: '{hoja}' de {hojas}")
    return df, errores


def cargar_inventario(archivo):
    """Lee el Excel de inventario y devuelve (df, errores)."""
    return _cargar(archivo, INV_COLUMNAS, FILA_ENCABEZADO_INVENTARIO, "Inventario")


def cargar_tiendas(archivo):
    """Lee el Excel de tiendas y devuelve (df, errores)."""
    return _cargar(
        archivo,
        TDAS_COLUMNAS,
        FILA_ENCABEZADO_TIENDAS,
        "Tiendas",
        TDAS_COLUMNAS_OPCIONALES,
    )
