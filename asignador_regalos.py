import io
from collections import deque
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Formato de fecha que entrega el sistema de origen. Se intenta primero y,
# si no calza, se recurre a un parseo genérico (ver `parsear_fechas`).
FORMATO_FECHA_ORIGEN = "%m/%d/%Y %I:%M:%S %p"

# Columnas que se usan para cruzar inventario con tiendas. Se normalizan a
# mayúsculas sin espacios en columnas auxiliares para que el cruce no falle
# por diferencias de captura ("lima" vs "LIMA").
COL_ZONA_KEY = "_ZonaKey"
COL_TIPO_KEY = "_TipoKey"

COLUMNAS_ASIGNACION = ["REGALO_1", "DESC_REGALO_1", "REGALO_2", "DESC_REGALO_2", "NOTAS"]

ANCHO_MINIMO_COLUMNA = 10
ANCHO_MAXIMO_COLUMNA = 60


# ============================
# Funciones auxiliares
# ============================
def normalizar_texto(df):
    """Recorta espacios de las columnas de texto sin convertir nulos en cadenas.

    `astype(str)` transformaría NaN en el literal "nan", que luego contamina
    los cruces y el Excel de salida. Aquí los nulos quedan como cadena vacía.
    """
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = df[col].astype("string").str.strip().fillna("")
    return df


def clave_normalizada(serie):
    """Genera la clave de cruce: sin espacios y en mayúsculas."""
    return serie.astype("string").str.strip().str.upper().fillna("")


def parsear_fechas(serie):
    """Convierte a datetime tolerando el formato de origen y otras variantes.

    Devuelve (serie_convertida, cantidad_no_reconocida). Fijar un único formato
    rígido hacía que un cambio de formato en el Excel volviera todo NaT en
    silencio, dejando sin sentido el orden de las estrategias por fecha.
    """
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie, 0

    convertida = pd.to_datetime(serie, errors="coerce", format=FORMATO_FECHA_ORIGEN)
    pendientes = convertida.isna() & serie.notna()
    if pendientes.any():
        try:
            alternativa = pd.to_datetime(
                serie[pendientes], errors="coerce", format="mixed"
            )
        except (ValueError, TypeError):
            alternativa = pd.to_datetime(serie[pendientes], errors="coerce")
        convertida.loc[pendientes] = alternativa

    no_reconocidas = int((convertida.isna() & serie.notna()).sum())
    return convertida, no_reconocidas


def ordenar_por_estrategia(df_inv, estrategia):
    """Ordena el inventario según la estrategia, preservando el índice original.

    El índice es la identidad de cada fila de stock: se usa después para
    devolver las cantidades descontadas al inventario global.
    """
    if estrategia == "Sobrantes":
        return df_inv.sort_values(
            by=[COL_TIPO_KEY, "FechaIngreso", "CantidadDisponible"],
            ascending=[True, True, True],
        )
    if estrategia == "Novedades":
        return df_inv.sort_values(
            by=[COL_TIPO_KEY, "FechaIngreso", "CantidadDisponible"],
            ascending=[True, False, False],
        )
    if estrategia == "AltoStock":
        return df_inv.sort_values(
            by=[COL_TIPO_KEY, "CantidadDisponible", "FechaIngreso"],
            ascending=[True, False, True],
        )
    if estrategia == "Equitativo":
        base = df_inv.sort_values(by=[COL_TIPO_KEY, "CodigoArticulo"])
        orden = []
        for _, grupo in base.groupby(COL_TIPO_KEY, sort=False):
            dq = deque(grupo.index)
            while dq:
                orden.append(dq.popleft())
                if dq:
                    dq.rotate(-1)
        return df_inv.loc[orden]
    return df_inv


def _tomar(df_inv_tipo, tomas):
    """Descuenta unidades y devuelve (ok, codigos, descripciones, inventario)."""
    inv = df_inv_tipo.copy()
    codigos, descripciones = [], []
    for idx, unidades in tomas:
        inv.loc[idx, "CantidadDisponible"] -= unidades
        codigos.extend([inv.at[idx, "CodigoArticulo"]] * unidades)
        descripciones.extend([inv.at[idx, "DescripcionArticulo"]] * unidades)
    return True, codigos, descripciones, inv


def intentar_asignar_para_tienda(df_inv_tipo, n_regalos):
    """Intenta asignar `n_regalos` de un inventario de un tipo específico.

    Para dos regalos se prefiere dar variedad a la tienda: primero dos
    artículos distintos y, sólo si no hay, dos unidades del mismo artículo.

    Devuelve (éxito, códigos, descripciones, inventario_actualizado).
    """
    inv = df_inv_tipo
    disponibles = inv.index[inv["CantidadDisponible"] >= 1]

    if n_regalos == 1:
        if len(disponibles) > 0:
            return _tomar(inv, [(disponibles[0], 1)])
        return False, [], [], df_inv_tipo

    if n_regalos == 2:
        # 1) Dos artículos distintos, una unidad de cada uno.
        if len(disponibles) >= 2:
            idx1 = disponibles[0]
            cod1 = inv.at[idx1, "CodigoArticulo"]
            for idx2 in disponibles[1:]:
                if inv.at[idx2, "CodigoArticulo"] != cod1:
                    return _tomar(inv, [(idx1, 1), (idx2, 1)])

        # 2) Dos unidades del mismo artículo.
        con_stock_doble = inv.index[inv["CantidadDisponible"] >= 2]
        if len(con_stock_doble) > 0:
            return _tomar(inv, [(con_stock_doble[0], 2)])

        # 3) Dos filas de stock del mismo artículo, una unidad de cada una.
        if len(disponibles) >= 2:
            return _tomar(inv, [(disponibles[0], 1), (disponibles[1], 1)])

    return False, [], [], df_inv_tipo


def _dar_formato_hoja(hoja):
    """Ajusta el ancho de las columnas y resalta la fila de encabezado."""
    for indice, columna in enumerate(hoja.columns, start=1):
        largo_maximo = 0
        for celda in columna:
            if celda.value is not None:
                largo_maximo = max(largo_maximo, len(str(celda.value)))
        ancho = min(max(largo_maximo + 2, ANCHO_MINIMO_COLUMNA), ANCHO_MAXIMO_COLUMNA)
        hoja.column_dimensions[get_column_letter(indice)].width = ancho

    relleno = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    fuente = Font(color="FFFFFF", bold=True)
    for celda in hoja[1]:
        celda.fill = relleno
        celda.font = fuente
    hoja.freeze_panes = "A2"


# ============================
# Función principal
# ============================
def ejecutar_asignacion(inv, tdas, n_regalos, estrategia):
    """Orquesta la asignación de regalos a tiendas con lógica de fallback.

    No modifica los DataFrames recibidos: trabaja siempre sobre copias.
    """
    # 1. Preparación de datos (sobre copias, para no mutar la entrada)
    inv = normalizar_texto(inv.copy())
    tdas = normalizar_texto(tdas.copy())

    inv["FechaIngreso"], fechas_no_reconocidas = parsear_fechas(inv["FechaIngreso"])
    inv["CantidadDisponible"] = (
        pd.to_numeric(inv["CantidadDisponible"], errors="coerce")
        .fillna(0)
        .astype("Int64")
    )
    inv = inv[inv["CantidadDisponible"] > 0].copy()

    # Claves auxiliares de cruce: insensibles a mayúsculas y espacios.
    inv[COL_ZONA_KEY] = clave_normalizada(inv["ZonaElegible"])
    inv[COL_TIPO_KEY] = clave_normalizada(inv["TipoRegalo"])
    tdas[COL_ZONA_KEY] = clave_normalizada(tdas["Zona"])
    tdas[COL_TIPO_KEY] = clave_normalizada(tdas["TipoRegalo"])

    # 2. Inicializar las columnas de asignación (texto libre)
    for columna in COLUMNAS_ASIGNACION:
        tdas[columna] = pd.Series("", index=tdas.index, dtype=object)

    excepciones = []
    parciales = 0
    inv_actualizado = inv.copy()

    def registrar_asignacion(idx_tienda, codigos, descripciones, nota=""):
        tdas.loc[idx_tienda, "REGALO_1"] = codigos[0]
        tdas.loc[idx_tienda, "DESC_REGALO_1"] = descripciones[0]
        if len(codigos) > 1:
            tdas.loc[idx_tienda, "REGALO_2"] = codigos[1]
            tdas.loc[idx_tienda, "DESC_REGALO_2"] = descripciones[1]
        tdas.loc[idx_tienda, "NOTAS"] = nota

    # 3. Iterar por cada zona
    for zona_key in sorted(tdas[COL_ZONA_KEY].unique()):
        tiendas_z = tdas[tdas[COL_ZONA_KEY] == zona_key]
        inv_z = inv_actualizado[inv_actualizado[COL_ZONA_KEY] == zona_key]

        if inv_z.empty:
            for idx_tienda, rowt in tiendas_z.iterrows():
                motivo = f"No hay inventario disponible en la zona {rowt['Zona']}"
                tdas.loc[idx_tienda, "NOTAS"] = motivo
                excepciones.append(
                    {
                        "IDTienda": rowt["IDTienda"],
                        "NombreTienda": rowt["NombreTienda"],
                        "Zona": rowt["Zona"],
                        "Motivo": motivo,
                    }
                )
            continue

        inv_z_ord = ordenar_por_estrategia(inv_z, estrategia)
        inv_por_tipo = {
            tipo: df.copy() for tipo, df in inv_z_ord.groupby(COL_TIPO_KEY, sort=False)
        }

        # 4. Iterar por cada tienda de la zona
        for idx_tienda, rowt in tiendas_z.iterrows():
            tipo_key = rowt[COL_TIPO_KEY]
            asignada = False

            if tipo_key in inv_por_tipo:
                ok, cods, descs, df_nuevo = intentar_asignar_para_tienda(
                    inv_por_tipo[tipo_key], n_regalos
                )
                if ok:
                    inv_por_tipo[tipo_key] = df_nuevo
                    registrar_asignacion(idx_tienda, cods, descs)
                    asignada = True

                # Fallback: entregar al menos un regalo.
                elif n_regalos > 1:
                    ok, cods, descs, df_nuevo = intentar_asignar_para_tienda(
                        inv_por_tipo[tipo_key], 1
                    )
                    if ok:
                        inv_por_tipo[tipo_key] = df_nuevo
                        registrar_asignacion(
                            idx_tienda,
                            cods,
                            descs,
                            f"Asignación parcial (1 de {n_regalos} solicitados)",
                        )
                        asignada = True
                        parciales += 1

            if not asignada:
                if tipo_key not in inv_por_tipo:
                    motivo = (
                        f"No hay inventario del tipo '{rowt['TipoRegalo']}' "
                        f"en la zona {rowt['Zona']}"
                    )
                else:
                    motivo = (
                        f"Stock insuficiente para asignar los {n_regalos} "
                        f"regalos solicitados"
                    )
                tdas.loc[idx_tienda, "NOTAS"] = motivo
                excepciones.append(
                    {
                        "IDTienda": rowt["IDTienda"],
                        "NombreTienda": rowt["NombreTienda"],
                        "Zona": rowt["Zona"],
                        "Motivo": motivo,
                    }
                )

        # 5. Consolidar los descuentos de la zona en el inventario global.
        #    La asignación por índice es segura porque `ordenar_por_estrategia`
        #    preserva las etiquetas originales de fila.
        for df_tipo in inv_por_tipo.values():
            inv_actualizado.loc[df_tipo.index, "CantidadDisponible"] = df_tipo[
                "CantidadDisponible"
            ]

    # 6. Preparar salida (sin las columnas auxiliares de cruce)
    auxiliares = [COL_ZONA_KEY, COL_TIPO_KEY]
    df_tiendas_final = tdas.drop(columns=auxiliares)
    df_inv_restante = (
        inv_actualizado[inv_actualizado["CantidadDisponible"] > 0]
        .drop(columns=auxiliares)
        .copy()
    )

    tiendas_asignadas = int(df_tiendas_final["REGALO_1"].ne("").sum())
    total_regalos = tiendas_asignadas + int(df_tiendas_final["REGALO_2"].ne("").sum())

    reporte = [
        "==== REPORTE DE EJECUCIÓN ====",
        f"Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"EstrategiaDePriorizacion: {estrategia}",
        f"NumeroRegalosPorTienda: {n_regalos}",
        f"Tiendas procesadas: {len(df_tiendas_final)}",
        f"Tiendas con asignación: {tiendas_asignadas}",
        f"Tiendas con asignación parcial: {parciales}",
        f"Total de regalos asignados: {total_regalos}",
        f"Unidades restantes en inventario: {int(df_inv_restante['CantidadDisponible'].sum())}",
    ]
    if fechas_no_reconocidas:
        reporte.append(
            f"ADVERTENCIA: {fechas_no_reconocidas} fecha(s) de ingreso no se "
            "pudieron interpretar; el orden por fecha puede ser inexacto."
        )
    reporte.append("\n---- Excepciones ----")
    if excepciones:
        for e in excepciones:
            reporte.append(
                f"[{e['Zona']}] {e['IDTienda']} - {e['NombreTienda']}: {e['Motivo']}"
            )
    else:
        reporte.append("Sin excepciones.")

    # 7. Exportar a Excel en memoria con formato
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_tiendas_final.to_excel(writer, index=False, sheet_name="Asignacion")
        df_inv_restante.to_excel(writer, index=False, sheet_name="InventarioRestante")
        for nombre_hoja in ("Asignacion", "InventarioRestante"):
            _dar_formato_hoja(writer.sheets[nombre_hoja])

    return df_tiendas_final, df_inv_restante, "\n".join(reporte), output.getvalue()
