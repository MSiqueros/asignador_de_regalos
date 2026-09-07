"""Pruebas de la lectura y validación de los Excel de entrada."""
import io
import sys
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from carga_datos import (
    INV_MAPPING,
    TDAS_MAPPING,
    cargar_inventario,
    cargar_tiendas,
    preparar_dataframe,
)


def construir_xlsx(encabezados, filas, filas_basura=0):
    """Crea un .xlsx en memoria, opcionalmente con filas basura arriba."""
    wb = Workbook()
    hoja = wb.active
    for _ in range(filas_basura):
        hoja.append(["Reporte generado por el sistema"])
    hoja.append(encabezados)
    for fila in filas:
        hoja.append(fila)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# preparar_dataframe
# ---------------------------------------------------------------------------
def test_renombra_las_columnas_al_nombre_interno():
    df = pd.DataFrame([{origen: "x" for origen in INV_MAPPING}])

    resultado, errores = preparar_dataframe(df, INV_MAPPING, "Inventario")

    assert errores == []
    assert set(INV_MAPPING.values()).issubset(resultado.columns)


def test_reporta_las_columnas_faltantes_sin_lanzar_excepcion():
    df = pd.DataFrame([{"CODIGO": 1}])

    resultado, errores = preparar_dataframe(df, TDAS_MAPPING, "Tiendas")

    assert resultado is None
    assert errores
    assert "NombreTienda" in errores[0]


def test_rechaza_archivos_con_columnas_duplicadas():
    df = pd.DataFrame([[1, 2]], columns=["CODIGO", "CODIGO"])

    resultado, errores = preparar_dataframe(df, TDAS_MAPPING, "Tiendas")

    assert resultado is None
    assert "duplicadas" in errores[0].lower()


def test_no_duplica_columnas_si_el_destino_ya_existe():
    """Si el Excel ya trae 'Zona', renombrar 'TERRITORIO' crearía un duplicado."""
    df = pd.DataFrame(
        [
            {
                "CODIGO": 1,
                "NOMBRE_COLABORADOR": "T1",
                "TERRITORIO": "A",
                "Zona": "A",
                "TIPOREGALO": "TIPO1",
            }
        ]
    )

    resultado, errores = preparar_dataframe(df, TDAS_MAPPING, "Tiendas")

    assert errores == []
    assert not resultado.columns.duplicated().any()


# ---------------------------------------------------------------------------
# Lectura de archivos reales
# ---------------------------------------------------------------------------
def test_el_inventario_se_lee_saltando_las_filas_de_encabezado():
    archivo = construir_xlsx(
        list(INV_MAPPING.keys()),
        [["01/15/2025 10:30:00 AM", "LIMA", "TIPO1", "ART-1", "Taza", 5]],
        filas_basura=2,
    )

    inv, errores = cargar_inventario(archivo)

    assert errores == []
    assert len(inv) == 1
    assert inv.loc[0, "CodigoArticulo"] == "ART-1"


def test_las_tiendas_se_leen_desde_la_primera_fila():
    archivo = construir_xlsx(list(TDAS_MAPPING.keys()), [[1, "Tienda A", "LIMA", "TIPO1"]])

    tdas, errores = cargar_tiendas(archivo)

    assert errores == []
    assert tdas.loc[0, "NombreTienda"] == "Tienda A"


def test_flujo_completo_desde_los_excel_hasta_la_asignacion():
    from asignador_regalos import ejecutar_asignacion

    inv_file = construir_xlsx(
        list(INV_MAPPING.keys()),
        [
            ["01/15/2025 10:30:00 AM", "LIMA", "TIPO1", "ART-1", "Taza roja", 3],
            ["01/16/2025 10:30:00 AM", "LIMA", "TIPO1", "ART-2", "Polo azul", 3],
        ],
        filas_basura=2,
    )
    tdas_file = construir_xlsx(
        list(TDAS_MAPPING.keys()),
        [[101, "Tienda Centro", "LIMA", "TIPO1"], [102, "Tienda Norte", "LIMA", "TIPO1"]],
    )

    inv, err_inv = cargar_inventario(inv_file)
    tdas, err_tdas = cargar_tiendas(tdas_file)
    assert err_inv == [] and err_tdas == []

    asignaciones, inv_rest, reporte, excel_bytes = ejecutar_asignacion(
        inv, tdas, 2, "Sobrantes"
    )

    assert asignaciones["REGALO_1"].ne("").all()
    assert asignaciones["REGALO_2"].ne("").all()
    # 2 tiendas x 2 regalos = 4 unidades entregadas de 6 disponibles
    assert int(inv_rest["CantidadDisponible"].sum()) == 2
    assert "Sin excepciones." in reporte
    assert excel_bytes[:2] == b"PK"  # un .xlsx es un zip
