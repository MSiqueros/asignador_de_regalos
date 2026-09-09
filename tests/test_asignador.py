"""Pruebas de la lógica de asignación de regalos.

Cada prueba documenta un comportamiento esperado del motor de asignación.
Se usan DataFrames construidos a mano con los nombres de columna que produce
el mapeo de `app.py` (ya renombrados).
"""
import io
import sys
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from asignador_regalos import ejecutar_asignacion, tomar_regalo

ESTRATEGIAS = ["Sobrantes", "Novedades", "AltoStock", "Equitativo"]


def construir_inventario(filas):
    """filas: lista de tuplas (zona, tipo, codigo, descripcion, cantidad)."""
    return pd.DataFrame(
        [
            {
                "FechaIngreso": "01/15/2025 10:30:00 AM",
                "ZonaElegible": zona,
                "TipoRegalo": tipo,
                "CodigoArticulo": codigo,
                "DescripcionArticulo": desc,
                "CantidadDisponible": cant,
            }
            for zona, tipo, codigo, desc, cant in filas
        ]
    )


def construir_tiendas(filas):
    """filas: (id, nombre, zona, tipo) o (id, nombre, zona, tipo, tipo_adicional)."""
    return pd.DataFrame(
        [
            {
                "IDTienda": idt,
                "NombreTienda": nom,
                "Zona": zona,
                "TipoRegalo": tipo,
                "TipoRegaloAdicional": adicional,
            }
            for idt, nom, zona, tipo, adicional in (
                fila if len(fila) == 5 else (*fila, "") for fila in filas
            )
        ]
    )


def regalos_asignados(df_tiendas):
    """Cuenta cuántas unidades se entregaron en total."""
    return int(df_tiendas["REGALO_1"].ne("").sum() + df_tiendas["REGALO_2"].ne("").sum())


# ---------------------------------------------------------------------------
# Conservación de stock: lo entregado + lo restante == lo inicial
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("estrategia", ESTRATEGIAS)
def test_el_stock_se_conserva_en_todas_las_estrategias(estrategia):
    inv = construir_inventario(
        [
            ("A", "TIPO1", "ART-A1", "Taza A1", 10),
            ("A", "TIPO1", "ART-A2", "Taza A2", 10),
            ("A", "TIPO1", "ART-A3", "Taza A3", 10),
            ("B", "TIPO1", "ART-B1", "Polo B1", 100),
            ("B", "TIPO1", "ART-B2", "Polo B2", 100),
            ("B", "TIPO1", "ART-B3", "Polo B3", 100),
        ]
    )
    tdas = construir_tiendas(
        [
            (1, "Tienda A", "A", "TIPO1"),
            (2, "Tienda B", "B", "TIPO1"),
        ]
    )
    total_inicial = 330

    asignaciones, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, estrategia)

    entregado = regalos_asignados(asignaciones)
    restante = int(inv_rest["CantidadDisponible"].sum())
    assert entregado == 2
    assert restante + entregado == total_inicial


@pytest.mark.parametrize("estrategia", ESTRATEGIAS)
def test_el_inventario_restante_no_contiene_filas_fantasma(estrategia):
    inv = construir_inventario(
        [
            ("A", "TIPO1", "ART-A1", "Taza A1", 5),
            ("B", "TIPO1", "ART-B1", "Polo B1", 5),
            ("B", "TIPO1", "ART-B2", "Polo B2", 5),
        ]
    )
    tdas = construir_tiendas([(1, "Tienda B", "B", "TIPO1")])

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, estrategia)

    assert inv_rest["ZonaElegible"].notna().all()
    assert inv_rest["CodigoArticulo"].notna().all()


# ---------------------------------------------------------------------------
# Parseo de fechas
# ---------------------------------------------------------------------------
def test_acepta_fechas_en_formato_iso():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    inv["FechaIngreso"] = ["2025-08-28 00:00:00"]
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert inv_rest["FechaIngreso"].notna().all()


def test_acepta_fechas_ya_tipadas_como_datetime():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    inv["FechaIngreso"] = pd.to_datetime(["2025-08-28"])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert inv_rest["FechaIngreso"].notna().all()


@pytest.mark.parametrize("vacia", [None, "", "   ", pd.NaT])
def test_las_fechas_ausentes_no_se_reportan_como_formato_no_reconocido(vacia):
    """El inventario deja la fecha en blanco para el stock de almacén.

    Una fecha ausente no es una fecha mal escrita: contarla como formato
    inválido dispara una advertencia falsa en cada ejecución.
    """
    inv = construir_inventario(
        [("A", "TIPO1", "ART-1", "Taza", 5), ("A", "TIPO1", "ART-2", "Polo", 5)]
    )
    inv["FechaIngreso"] = ["01/15/2025 10:30:00 AM", vacia]
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    _, _, reporte, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert "ADVERTENCIA" not in reporte


def test_sigue_advirtiendo_cuando_la_fecha_es_ilegible():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    inv["FechaIngreso"] = ["no es una fecha"]
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    _, _, reporte, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert "ADVERTENCIA" in reporte
    assert "1 fecha(s)" in reporte


# ---------------------------------------------------------------------------
# Normalización de texto
# ---------------------------------------------------------------------------
def test_los_nulos_de_texto_no_se_convierten_en_la_cadena_nan():
    """read_excel entrega las celdas vacías como NaN; no deben llegar como texto."""
    inv = construir_inventario(
        [
            ("A", "TIPO1", "ART-1", float("nan"), 5),
            ("A", "TIPO1", "ART-2", "Polo azul", 5),
            ("A", "TIPO1", "ART-3", None, 5),
        ]
    )
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    descripciones = inv_rest["DescripcionArticulo"].tolist()
    assert "nan" not in descripciones
    assert "None" not in descripciones
    assert "NaT" not in descripciones


def test_el_cruce_de_zona_y_tipo_ignora_mayusculas_y_espacios():
    inv = construir_inventario([("  lima ", "tipo1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "LIMA", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"


def test_el_cruce_de_zona_y_tipo_ignora_los_acentos():
    """El inventario escribe 'Huarochirí' y la plantilla de tiendas 'HUAROCHIRI'.

    Sin normalizar acentos la zona entera queda sin asignación.
    """
    inv = construir_inventario([("Huarochirí", "PEQUEÑO", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "HUAROCHIRI", "PEQUENO")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"


# ---------------------------------------------------------------------------
# Reglas de asignación
# ---------------------------------------------------------------------------
def test_con_dos_regalos_prefiere_articulos_distintos():
    inv = construir_inventario(
        [
            ("A", "TIPO1", "ART-1", "Taza", 5),
            ("A", "TIPO1", "ART-2", "Polo", 5),
        ]
    )
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] != asignaciones.loc[0, "REGALO_2"]


def test_con_dos_regalos_repite_articulo_si_no_hay_otro_distinto():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[0, "REGALO_2"] == "ART-1"


def test_asignacion_parcial_cuando_solo_queda_una_unidad():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 1)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[0, "REGALO_2"] == ""
    assert "parcial" in asignaciones.loc[0, "NOTAS"].lower()


# ---------------------------------------------------------------------------
# Salida
# ---------------------------------------------------------------------------
def test_la_salida_incluye_las_descripciones_de_los_regalos():
    inv = construir_inventario(
        [
            ("A", "TIPO1", "ART-1", "Taza roja", 5),
            ("A", "TIPO1", "ART-2", "Polo azul", 5),
        ]
    )
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert "DESC_REGALO_1" in asignaciones.columns
    assert "DESC_REGALO_2" in asignaciones.columns
    assert asignaciones.loc[0, "DESC_REGALO_1"] in ("Taza roja", "Polo azul")


def test_las_notas_explican_por_que_una_tienda_no_recibio_regalo():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(9, "Tienda Sin Stock", "Z", "TIPO1")])

    asignaciones, _, reporte, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == ""
    assert asignaciones.loc[0, "NOTAS"] != ""
    assert "Tienda Sin Stock" in reporte


def test_el_ancho_de_columna_considera_las_celdas_numericas():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1234567890123456789, "Tienda A", "A", "TIPO1")])

    _, _, _, excel_bytes = ejecutar_asignacion(inv, tdas, "Sobrantes")

    hoja = load_workbook(io.BytesIO(excel_bytes))["Asignacion"]
    columna_id = next(
        c[0].column_letter for c in hoja.columns if c[0].value == "IDTienda"
    )
    assert hoja.column_dimensions[columna_id].width >= 21


def test_no_muta_los_dataframes_recibidos():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])
    columnas_tdas_antes = list(tdas.columns)
    fecha_inv_antes = inv["FechaIngreso"].tolist()

    ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert list(tdas.columns) == columnas_tdas_antes
    assert inv["FechaIngreso"].tolist() == fecha_inv_antes


# ---------------------------------------------------------------------------
# tomar_regalo: entrega una unidad de un pozo de un solo tipo
# ---------------------------------------------------------------------------
def pozo(filas):
    """filas: lista de tuplas (codigo, descripcion, cantidad)."""
    return pd.DataFrame(
        [
            {
                "CodigoArticulo": codigo,
                "DescripcionArticulo": desc,
                "CantidadDisponible": cant,
            }
            for codigo, desc, cant in filas
        ]
    )


def test_tomar_regalo_entrega_una_unidad_y_la_descuenta():
    inv = pozo([("ART-1", "Taza", 3)])

    ok, codigo, desc, inv_nuevo = tomar_regalo(inv)

    assert ok
    assert codigo == "ART-1"
    assert desc == "Taza"
    assert int(inv_nuevo["CantidadDisponible"].sum()) == 2


def test_tomar_regalo_no_muta_el_pozo_recibido():
    inv = pozo([("ART-1", "Taza", 3)])

    tomar_regalo(inv)

    assert int(inv["CantidadDisponible"].sum()) == 3


def test_tomar_regalo_prefiere_un_articulo_no_excluido():
    inv = pozo([("ART-1", "Taza", 3), ("ART-2", "Polo", 3)])

    ok, codigo, _, _ = tomar_regalo(inv, codigos_excluidos=["ART-1"])

    assert ok
    assert codigo == "ART-2"


def test_tomar_regalo_repite_el_excluido_si_no_hay_alternativa():
    inv = pozo([("ART-1", "Taza", 3)])

    ok, codigo, _, _ = tomar_regalo(inv, codigos_excluidos=["ART-1"])

    assert ok
    assert codigo == "ART-1"


def test_tomar_regalo_avisa_cuando_el_pozo_esta_agotado():
    inv = pozo([("ART-1", "Taza", 0)])

    ok, codigo, desc, inv_nuevo = tomar_regalo(inv)

    assert not ok
    assert codigo == ""
    assert desc == ""
    assert int(inv_nuevo["CantidadDisponible"].sum()) == 0


# ---------------------------------------------------------------------------
# Regalo adicional y prioridad del primer regalo
# ---------------------------------------------------------------------------
def test_el_primer_regalo_de_toda_tienda_va_antes_que_un_adicional():
    """La regla central: nadie recibe el segundo mientras falte un primero.

    Con 2 unidades y 2 tiendas, la primera del archivo pide además un
    adicional. La lógica anterior le daba las dos unidades y dejaba a la
    segunda tienda sin nada.
    """
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 2)])
    tdas = construir_tiendas(
        [
            (1, "Tienda Primera", "A", "TIPO1", "TIPO1"),
            (2, "Tienda Segunda", "A", "TIPO1"),
        ]
    )

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[1, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[0, "REGALO_2"] == ""
    assert "parcial" in asignaciones.loc[0, "NOTAS"].lower()


def test_el_regalo_adicional_sale_del_pozo_de_su_propio_tipo():
    inv = construir_inventario(
        [
            ("A", "TIPO1", "ART-1", "Taza", 5),
            ("A", "TIPO2", "ART-2", "Polo", 5),
        ]
    )
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", "TIPO2")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[0, "REGALO_2"] == "ART-2"
    assert asignaciones.loc[0, "NOTAS"] == ""


@pytest.mark.parametrize("vacia", ["", "   ", None])
def test_sin_regalo_adicional_la_tienda_recibe_uno_solo(vacia):
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", vacia)])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[0, "REGALO_2"] == ""
    assert asignaciones.loc[0, "NOTAS"] == ""


def test_parcial_cuando_el_tipo_adicional_no_existe_en_la_zona():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", "TIPO9")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[0, "REGALO_2"] == ""
    assert "regalo adicional" in asignaciones.loc[0, "NOTAS"].lower()
    assert "TIPO9" in asignaciones.loc[0, "NOTAS"]


def test_si_falta_el_primer_regalo_el_adicional_ocupa_su_lugar():
    """REGALO_1 nunca queda vacío si hubo algo que entregar."""
    inv = construir_inventario([("A", "TIPO2", "ART-2", "Polo", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", "TIPO2")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-2"
    assert asignaciones.loc[0, "REGALO_2"] == ""
    assert "TIPO1" in asignaciones.loc[0, "NOTAS"]


def test_sin_stock_de_ninguno_de_los_dos_tipos_no_hay_asignacion():
    inv = construir_inventario([("A", "TIPO3", "ART-3", "Gorra", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1", "TIPO2")])

    asignaciones, _, reporte, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == ""
    assert asignaciones.loc[0, "REGALO_2"] == ""
    assert "Tienda A" in reporte


def test_el_motor_tolera_una_tienda_sin_la_columna_de_regalo_adicional():
    """El motor no puede exigir lo que la carga declara opcional."""
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")]).drop(
        columns=["TipoRegaloAdicional"]
    )

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[0, "REGALO_2"] == ""


def test_el_adicional_de_una_zona_no_consume_el_stock_de_otra():
    """Las dos pasadas son por zona; hacerlas globalmente daría lo mismo."""
    inv = construir_inventario(
        [
            ("A", "TIPO1", "ART-A", "Taza A", 1),
            ("B", "TIPO1", "ART-B", "Taza B", 5),
        ]
    )
    tdas = construir_tiendas(
        [
            (1, "Tienda A", "A", "TIPO1", "TIPO1"),
            (2, "Tienda B", "B", "TIPO1"),
        ]
    )

    asignaciones, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-A"
    assert asignaciones.loc[0, "REGALO_2"] == ""
    assert asignaciones.loc[1, "REGALO_1"] == "ART-B"
    # De la zona B solo salió el regalo de su propia tienda.
    restante_b = inv_rest[inv_rest["ZonaElegible"] == "B"]["CantidadDisponible"].sum()
    assert int(restante_b) == 4


def test_el_reporte_cuenta_los_regalos_adicionales():
    inv = construir_inventario(
        [
            ("A", "TIPO1", "ART-1", "Taza", 5),
            ("A", "TIPO2", "ART-2", "Polo", 1),
        ]
    )
    tdas = construir_tiendas(
        [
            (1, "Tienda A", "A", "TIPO1", "TIPO2"),
            (2, "Tienda B", "A", "TIPO1", "TIPO2"),
        ]
    )

    asignaciones, _, reporte, _ = ejecutar_asignacion(inv, tdas, "Sobrantes")

    assert "Tiendas que piden regalo adicional: 2" in reporte
    assert "Regalos adicionales entregados: 1" in reporte
    assert asignaciones.attrs["metricas"]["piden_adicional"] == 2
    assert asignaciones.attrs["metricas"]["adicionales_entregados"] == 1
