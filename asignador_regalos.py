import io
import unicodedata
from collections import deque
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from carga_datos import normalizar_encabezado

# Formato de fecha que entrega el sistema de origen. Se intenta primero y,
# si no calza, se recurre a un parseo genérico (ver `parsear_fechas`).
FORMATO_FECHA_ORIGEN = "%m/%d/%Y %I:%M:%S %p"

# Columnas que se usan para cruzar inventario con tiendas. Se normalizan a
# mayúsculas sin espacios en columnas auxiliares para que el cruce no falle
# por diferencias de captura ("lima" vs "LIMA").
COL_ZONA_KEY = "_ZonaKey"
COL_TIPO_KEY = "_TipoKey"
COL_TIPO_ADIC_KEY = "_TipoAdicKey"

COLUMNAS_ASIGNACION = ["REGALO_1", "DESC_REGALO_1", "REGALO_2", "DESC_REGALO_2", "NOTAS"]

# El export de inventario repite la misma cantidad en varias columnas.
# `preparar_dataframe` solo renombra la primera que encuentra, así que las demás
# llegarían a la salida con el valor original y contradirían al stock ya
# descontado: el usuario ve 'CantidadDisponible 5' junto a 'SALDO 39'.
# Se sincronizan por nombre normalizado, no por posición.
COLUMNAS_ESPEJO_STOCK = ("SALDO",)
COLUMNAS_ESPEJO_ENTREGADAS = ("CANTIDADENTREGADA", "CONTIDADENTREGADA")

# Columna propia, siempre presente: cuántas unidades salieron de cada fila de
# stock en esta corrida. No depende de que el archivo de origen traiga la suya.
COL_ENTREGADAS = "UnidadesEntregadas"

# Cuántas tiendas se listan por motivo antes de resumir el resto en un conteo.
MAX_EJEMPLOS_POR_MOTIVO = 5

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


def _sin_acentos(valor):
    descompuesto = unicodedata.normalize("NFKD", valor)
    return "".join(c for c in descompuesto if not unicodedata.combining(c))


def clave_normalizada(serie):
    """Genera la clave de cruce: sin espacios, sin acentos y en mayúsculas.

    Los acentos importan: el inventario escribe "Huarochirí" y la plantilla de
    tiendas "HUAROCHIRI". Sin normalizarlos la zona completa queda sin cruce y
    sus tiendas se reportan como "sin inventario" aunque haya stock.
    """
    return (
        serie.astype("string")
        .fillna("")
        .map(_sin_acentos)
        .astype("string")
        .str.strip()
        .str.upper()
    )


def parsear_fechas(serie):
    """Convierte a datetime tolerando el formato de origen y otras variantes.

    Devuelve (serie_convertida, cantidad_no_reconocida). Fijar un único formato
    rígido hacía que un cambio de formato en el Excel volviera todo NaT en
    silencio, dejando sin sentido el orden de las estrategias por fecha.
    """
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie, 0

    # Una fecha ausente no es una fecha mal escrita. El inventario deja la
    # fecha en blanco para parte del stock y `normalizar_texto` ya convirtió
    # esos nulos en cadena vacía, así que hay que excluir ambos casos: contarlos
    # dispararía una advertencia de "formato no reconocido" en cada ejecución.
    presente = serie.notna() & serie.astype("string").fillna("").str.strip().ne("")

    convertida = pd.to_datetime(serie, errors="coerce", format=FORMATO_FECHA_ORIGEN)
    pendientes = convertida.isna() & presente
    if pendientes.any():
        try:
            alternativa = pd.to_datetime(
                serie[pendientes], errors="coerce", format="mixed"
            )
        except (ValueError, TypeError):
            alternativa = pd.to_datetime(serie[pendientes], errors="coerce")
        convertida.loc[pendientes] = alternativa

    no_reconocidas = int((convertida.isna() & presente).sum())
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


def tomar_regalo(df_inv_tipo, codigos_excluidos=()):
    """Entrega una unidad del pozo recibido, prefiriendo un artículo nuevo.

    `codigos_excluidos` son los artículos que esa tienda ya recibió: se usan
    para darle variedad cuando su regalo adicional es del mismo tipo que el
    primero. Si no hay otro artículo con stock, se repite el mismo.

    Devuelve (ok, codigo, descripcion, inventario_actualizado). No muta el
    DataFrame recibido.
    """
    inv = df_inv_tipo
    disponibles = inv.index[inv["CantidadDisponible"] >= 1]
    if len(disponibles) == 0:
        return False, "", "", df_inv_tipo

    idx = next(
        (
            i
            for i in disponibles
            if inv.at[i, "CodigoArticulo"] not in codigos_excluidos
        ),
        disponibles[0],
    )

    inv = inv.copy()
    inv.loc[idx, "CantidadDisponible"] -= 1
    return True, inv.at[idx, "CodigoArticulo"], inv.at[idx, "DescripcionArticulo"], inv


def servir_de_pozo(inv_por_tipo, tipo_key, codigos_excluidos=()):
    """Toma una unidad del pozo `tipo_key` y actualiza el diccionario in situ.

    Un tipo que no existe en la zona se trata igual que un pozo agotado: no
    es un error, es una asignación que no se pudo completar.

    Devuelve (ok, codigo, descripcion).
    """
    pozo = inv_por_tipo.get(tipo_key)
    if pozo is None:
        return False, "", ""

    ok, codigo, descripcion, pozo_nuevo = tomar_regalo(pozo, codigos_excluidos)
    if ok:
        inv_por_tipo[tipo_key] = pozo_nuevo
    return ok, codigo, descripcion


def sincronizar_columnas_espejo(df_inv, entregadas):
    """Pone al día las columnas del origen que repiten la cantidad.

    `entregadas` es la serie de unidades que salieron de cada fila. Sin esto,
    columnas como 'SALDO' o 'CANTIDADENTREGADA' quedan con el valor previo a la
    asignación y contradicen a 'CantidadDisponible'.

    Las de entregadas se **suman**, no se pisan: si el archivo ya traía un
    conteo de un reparto anterior, esta corrida se acumula sobre él.
    """
    df = df_inv.copy()
    df[COL_ENTREGADAS] = entregadas.reindex(df.index).fillna(0).astype("Int64")

    for columna in df.columns:
        if columna == COL_ENTREGADAS:
            continue
        clave = normalizar_encabezado(columna)
        if clave in COLUMNAS_ESPEJO_STOCK:
            df[columna] = df["CantidadDisponible"]
        elif clave in COLUMNAS_ESPEJO_ENTREGADAS:
            previas = pd.to_numeric(df[columna], errors="coerce").fillna(0)
            df[columna] = (previas.astype("Int64") + df[COL_ENTREGADAS]).astype("Int64")
    return df


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
# Reporte de texto
# ============================
ANCHO_REPORTE = 66


def _titulo(texto):
    return ["=" * ANCHO_REPORTE, f"  {texto}", "=" * ANCHO_REPORTE]


def _seccion(texto):
    """Encabezado de sección con la regla completando el ancho fijo."""
    prefijo = f"---- {texto} "
    return ["", prefijo + "-" * max(ANCHO_REPORTE - len(prefijo), 0)]


def _dato(etiqueta, valor, total=None, sangria=2):
    """Línea 'etiqueta ....... valor', con porcentaje si se da un total."""
    porcentaje = ""
    if total:
        porcentaje = f"  ({100 * valor / total:5.1f}%)"
    izquierda = " " * sangria + etiqueta
    return f"{izquierda:<44}{valor:>8}{porcentaje}"


def _redactar_reporte(
    *,
    estrategia,
    tiendas,
    asignadas,
    parciales,
    piden_adicional,
    adicionales_entregados,
    total_regalos,
    restantes,
    fechas_no_reconocidas,
    excepciones,
):
    """Arma el reporte de texto legible de la corrida.

    Se agrupa por secciones y las excepciones se consolidan por motivo: con
    cientos de tiendas sin asignación, la lista plana era ilegible y escondía
    que casi siempre se trata de unos pocos motivos repetidos.
    """
    completas = asignadas - parciales
    sin_asignacion = tiendas - asignadas

    lineas = _titulo("ASIGNADOR DE REGALOS  ·  REPORTE DE EJECUCIÓN")
    lineas += [
        f"  Fecha de ejecución : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"  Estrategia         : {estrategia}",
    ]

    lineas += _seccion("COBERTURA DE TIENDAS")
    lineas += [
        _dato("Tiendas procesadas", tiendas),
        _dato("Con asignación completa", completas, tiendas),
        _dato("Con asignación parcial", parciales, tiendas),
        _dato("Sin asignación", sin_asignacion, tiendas),
    ]

    lineas += _seccion("REGALOS ENTREGADOS")
    lineas += [
        _dato("Total de regalos entregados", total_regalos),
        _dato("· primer regalo", asignadas),
        _dato("· regalo adicional", adicionales_entregados),
    ]
    if piden_adicional:
        lineas += [
            _dato("Tiendas que piden regalo adicional", piden_adicional),
            _dato(
                "Adicionales cubiertos",
                adicionales_entregados,
                piden_adicional,
            ),
        ]
    else:
        lineas.append(
            "  Ninguna tienda pidió regalo adicional "
            "(columna 'Regalo adicional' vacía)."
        )

    lineas += _seccion("INVENTARIO")
    lineas += [
        _dato("Unidades entregadas", total_regalos),
        _dato("Unidades restantes", restantes),
    ]

    if fechas_no_reconocidas:
        lineas += _seccion("ADVERTENCIAS")
        lineas.append(
            f"  {fechas_no_reconocidas} fecha(s) de ingreso no se pudieron "
            "interpretar;"
        )
        lineas.append("  el orden de las estrategias por fecha puede ser inexacto.")

    lineas += _seccion(f"EXCEPCIONES ({len(excepciones)})")
    if not excepciones:
        lineas.append("  Sin excepciones: todas las tiendas recibieron su regalo.")
    else:
        por_motivo = {}
        for e in excepciones:
            por_motivo.setdefault(e["Motivo"], []).append(e)

        lineas.append(f"  {len(por_motivo)} motivo(s) distinto(s), del más frecuente:")
        for motivo, casos in sorted(
            por_motivo.items(), key=lambda kv: len(kv[1]), reverse=True
        ):
            lineas += ["", f"  [{len(casos):>4}]  {motivo}"]
            for e in casos[:MAX_EJEMPLOS_POR_MOTIVO]:
                lineas.append(
                    f"          · {e['IDTienda']}  {e['NombreTienda']}  ({e['Zona']})"
                )
            if len(casos) > MAX_EJEMPLOS_POR_MOTIVO:
                lineas.append(
                    f"          … y {len(casos) - MAX_EJEMPLOS_POR_MOTIVO} tienda(s) "
                    "más con este mismo motivo."
                )

        lineas += [
            "",
            "  El detalle completo, tienda por tienda, está en la columna NOTAS",
            "  de la hoja 'Asignacion'.",
        ]

    return "\n".join(lineas)


# ============================
# Función principal
# ============================
def ejecutar_asignacion(inv, tdas, estrategia):
    """Orquesta la asignación de regalos a tiendas en dos pasadas por zona.

    El primer regalo de todas las tiendas de una zona tiene prioridad sobre
    cualquier regalo adicional: recién cuando todas tuvieron su oportunidad se
    reparte lo que sobró entre las que piden un segundo obsequio.

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
    # La columna es opcional en la plantilla: si no llegó, nadie pide adicional.
    if "TipoRegaloAdicional" in tdas.columns:
        tdas[COL_TIPO_ADIC_KEY] = clave_normalizada(tdas["TipoRegaloAdicional"])
    else:
        tdas[COL_TIPO_ADIC_KEY] = ""

    # 2. Inicializar las columnas de asignación (texto libre)
    for columna in COLUMNAS_ASIGNACION:
        tdas[columna] = pd.Series("", index=tdas.index, dtype=object)

    excepciones = []
    parciales = 0
    piden_adicional = 0
    adicionales_entregados = 0
    inv_actualizado = inv.copy()

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

        # Lo que cada tienda de la zona consiguió y lo que le faltó.
        entregados = {idx: [] for idx in tiendas_z.index}
        faltantes = {idx: [] for idx in tiendas_z.index}

        # 4. Pasada 1: el primer regalo de todas las tiendas de la zona.
        for idx_tienda, rowt in tiendas_z.iterrows():
            ok, cod, desc = servir_de_pozo(inv_por_tipo, rowt[COL_TIPO_KEY])
            if ok:
                entregados[idx_tienda].append((cod, desc))
            else:
                faltantes[idx_tienda].append(("principal", rowt["TipoRegalo"]))

        # 5. Pasada 2: recién ahora los adicionales, con el stock sobrante.
        for idx_tienda, rowt in tiendas_z.iterrows():
            if not rowt[COL_TIPO_ADIC_KEY]:
                continue
            piden_adicional += 1
            ya_dados = [codigo for codigo, _ in entregados[idx_tienda]]
            ok, cod, desc = servir_de_pozo(
                inv_por_tipo, rowt[COL_TIPO_ADIC_KEY], ya_dados
            )
            if ok:
                entregados[idx_tienda].append((cod, desc))
                adicionales_entregados += 1
            else:
                faltantes[idx_tienda].append(
                    ("adicional", rowt["TipoRegaloAdicional"])
                )

        # 6. Volcar el resultado de la zona a las columnas de salida.
        for idx_tienda, rowt in tiendas_z.iterrows():
            recibidos = entregados[idx_tienda]
            sin_stock = faltantes[idx_tienda]

            if not recibidos:
                # Si el tipo principal y el adicional coinciden, nombrarlo dos
                # veces ("'mayorista' ni 'mayorista'") no aporta nada.
                distintas = list(dict.fromkeys(etiqueta for _, etiqueta in sin_stock))
                etiquetas = " ni ".join(f"'{etiqueta}'" for etiqueta in distintas)
                motivo = (
                    f"No hay inventario del tipo {etiquetas} "
                    f"en la zona {rowt['Zona']}"
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
                continue

            # Los regalos conseguidos ocupan las ranuras en orden, sin huecos:
            # si faltó el principal, el adicional queda en REGALO_1.
            for ranura, (codigo, descripcion) in enumerate(recibidos, start=1):
                tdas.loc[idx_tienda, f"REGALO_{ranura}"] = codigo
                tdas.loc[idx_tienda, f"DESC_REGALO_{ranura}"] = descripcion

            if sin_stock:
                clase, etiqueta = sin_stock[0]
                que = "del regalo adicional" if clase == "adicional" else "del tipo"
                tdas.loc[idx_tienda, "NOTAS"] = (
                    f"Asignación parcial: sin stock {que} '{etiqueta}' "
                    f"en la zona {rowt['Zona']}"
                )
                parciales += 1

        # 7. Consolidar los descuentos de la zona en el inventario global.
        #    La asignación por índice es segura porque `ordenar_por_estrategia`
        #    preserva las etiquetas originales de fila.
        for df_tipo in inv_por_tipo.values():
            inv_actualizado.loc[df_tipo.index, "CantidadDisponible"] = df_tipo[
                "CantidadDisponible"
            ]

    # 8. Preparar salida (sin las columnas auxiliares de cruce). La clave del
    #    tipo adicional solo existe en tiendas, no en el inventario.
    df_tiendas_final = tdas.drop(
        columns=[COL_ZONA_KEY, COL_TIPO_KEY, COL_TIPO_ADIC_KEY]
    )
    # Unidades que salieron de cada fila de stock en esta corrida.
    entregadas_por_fila = inv["CantidadDisponible"] - inv_actualizado[
        "CantidadDisponible"
    ]
    df_inv_restante = sincronizar_columnas_espejo(
        inv_actualizado[inv_actualizado["CantidadDisponible"] > 0]
        .drop(columns=[COL_ZONA_KEY, COL_TIPO_KEY])
        .copy(),
        entregadas_por_fila,
    )

    tiendas_asignadas = int(df_tiendas_final["REGALO_1"].ne("").sum())
    total_regalos = tiendas_asignadas + int(df_tiendas_final["REGALO_2"].ne("").sum())

    # "Regalos adicionales entregados" no se puede derivar de REGALO_2: cuando
    # falta el principal, el adicional ocupa REGALO_1 y ese conteo lo perdería.
    df_tiendas_final.attrs["metricas"] = {
        "piden_adicional": piden_adicional,
        "adicionales_entregados": adicionales_entregados,
        "parciales": parciales,
    }

    reporte = _redactar_reporte(
        estrategia=estrategia,
        tiendas=len(df_tiendas_final),
        asignadas=tiendas_asignadas,
        parciales=parciales,
        piden_adicional=piden_adicional,
        adicionales_entregados=adicionales_entregados,
        total_regalos=total_regalos,
        restantes=int(df_inv_restante["CantidadDisponible"].sum()),
        fechas_no_reconocidas=fechas_no_reconocidas,
        excepciones=excepciones,
    )

    # 9. Exportar a Excel en memoria con formato
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_tiendas_final.to_excel(writer, index=False, sheet_name="Asignacion")
        df_inv_restante.to_excel(writer, index=False, sheet_name="InventarioRestante")
        for nombre_hoja in ("Asignacion", "InventarioRestante"):
            _dar_formato_hoja(writer.sheets[nombre_hoja])

    return df_tiendas_final, df_inv_restante, reporte, output.getvalue()
