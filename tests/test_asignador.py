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

from asignador_regalos import ejecutar_asignacion

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
    """filas: lista de tuplas (id_tienda, nombre, zona, tipo)."""
    return pd.DataFrame(
        [
            {"IDTienda": idt, "NombreTienda": nom, "Zona": zona, "TipoRegalo": tipo}
            for idt, nom, zona, tipo in filas
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

    asignaciones, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, 1, estrategia)

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

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, 1, estrategia)

    assert inv_rest["ZonaElegible"].notna().all()
    assert inv_rest["CodigoArticulo"].notna().all()


# ---------------------------------------------------------------------------
# Parseo de fechas
# ---------------------------------------------------------------------------
def test_acepta_fechas_en_formato_iso():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    inv["FechaIngreso"] = ["2025-08-28 00:00:00"]
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

    assert inv_rest["FechaIngreso"].notna().all()


def test_acepta_fechas_ya_tipadas_como_datetime():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    inv["FechaIngreso"] = pd.to_datetime(["2025-08-28"])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

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

    _, _, reporte, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

    assert "ADVERTENCIA" not in reporte


def test_sigue_advirtiendo_cuando_la_fecha_es_ilegible():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    inv["FechaIngreso"] = ["no es una fecha"]
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    _, _, reporte, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

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

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

    descripciones = inv_rest["DescripcionArticulo"].tolist()
    assert "nan" not in descripciones
    assert "None" not in descripciones
    assert "NaT" not in descripciones


def test_el_cruce_de_zona_y_tipo_ignora_mayusculas_y_espacios():
    inv = construir_inventario([("  lima ", "tipo1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "LIMA", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"


def test_el_cruce_de_zona_y_tipo_ignora_los_acentos():
    """El inventario escribe 'Huarochirí' y la plantilla de tiendas 'HUAROCHIRI'.

    Sin normalizar acentos la zona entera queda sin asignación.
    """
    inv = construir_inventario([("Huarochirí", "PEQUEÑO", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "HUAROCHIRI", "PEQUENO")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

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
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 2, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] != asignaciones.loc[0, "REGALO_2"]


def test_con_dos_regalos_repite_articulo_si_no_hay_otro_distinto():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 2, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == "ART-1"
    assert asignaciones.loc[0, "REGALO_2"] == "ART-1"


def test_asignacion_parcial_cuando_solo_queda_una_unidad():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 1)])
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 2, "Sobrantes")

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
    tdas = construir_tiendas([(1, "Tienda A", "A", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 2, "Sobrantes")

    assert "DESC_REGALO_1" in asignaciones.columns
    assert "DESC_REGALO_2" in asignaciones.columns
    assert asignaciones.loc[0, "DESC_REGALO_1"] in ("Taza roja", "Polo azul")


def test_las_notas_explican_por_que_una_tienda_no_recibio_regalo():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(9, "Tienda Sin Stock", "Z", "TIPO1")])

    asignaciones, _, reporte, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

    assert asignaciones.loc[0, "REGALO_1"] == ""
    assert asignaciones.loc[0, "NOTAS"] != ""
    assert "Tienda Sin Stock" in reporte


def test_el_ancho_de_columna_considera_las_celdas_numericas():
    inv = construir_inventario([("A", "TIPO1", "ART-1", "Taza", 5)])
    tdas = construir_tiendas([(1234567890123456789, "Tienda A", "A", "TIPO1")])

    _, _, _, excel_bytes = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

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

    ejecutar_asignacion(inv, tdas, 1, "Sobrantes")

    assert list(tdas.columns) == columnas_tdas_antes
    assert inv["FechaIngreso"].tolist() == fecha_inv_antes
