"""Pruebas de la lectura y validación de los Excel de entrada."""
import io
import sys
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from carga_datos import (
    INV_COLUMNAS,
    TDAS_COLUMNAS,
    cargar_inventario,
    cargar_tiendas,
    normalizar_encabezado,
    preparar_dataframe,
)


def encabezados_de(columnas):
    """Primer alias de cada columna: los nombres que trae el export real."""
    return [alias[0] for alias in columnas.values()]


def construir_xlsx(
    encabezados, filas, filas_basura=0, hoja_previa=None, nombre_hoja="Datos"
):
    """Crea un .xlsx en memoria.

    `filas_basura` simula las filas de título que el export pone sobre el
    encabezado. `hoja_previa` inserta antes una hoja señuelo, como la tabla
    dinámica que la plantilla real de tiendas trae en primer lugar.
    """
    wb = Workbook()
    hoja = wb.active
    if hoja_previa is not None:
        hoja.title = "Hoja2"
        for fila in hoja_previa:
            hoja.append(fila)
        hoja = wb.create_sheet(nombre_hoja)
    else:
        hoja.title = nombre_hoja
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
# normalizar_encabezado
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("tamaño", "TAMANO"),
        ("TAMAÑO", "TAMANO"),
        ("Tamano", "TAMANO"),
        ("NOMBRE_COLABORADOR", "NOMBRECOLABORADOR"),
        ("nombre colaborador", "NOMBRECOLABORADOR"),
        ("  CODIGO  ", "CODIGO"),
    ],
)
def test_normalizar_encabezado_ignora_acentos_espacios_y_separadores(entrada, esperado):
    assert normalizar_encabezado(entrada) == esperado


# ---------------------------------------------------------------------------
# preparar_dataframe
# ---------------------------------------------------------------------------
def test_renombra_las_columnas_al_nombre_interno():
    df = pd.DataFrame([{origen: "x" for origen in encabezados_de(INV_COLUMNAS)}])

    resultado, errores = preparar_dataframe(df, INV_COLUMNAS, "Inventario")

    assert errores == []
    assert set(INV_COLUMNAS).issubset(resultado.columns)


def test_reconoce_tamano_como_tipo_de_regalo():
    """La plantilla de tiendas llama 'tamaño' al segmento del inventario."""
    df = pd.DataFrame(
        [
            {
                "CODIGO": 1,
                "NOMBRE_COLABORADOR": "Tienda A",
                "TERRITORIO": "LIMA SUR",
                "tamaño": "mediana",
            }
        ]
    )

    resultado, errores = preparar_dataframe(df, TDAS_COLUMNAS, "Tiendas")

    assert errores == []
    assert resultado.loc[0, "TipoRegalo"] == "mediana"


def test_reconoce_los_encabezados_sin_importar_acentos_ni_mayusculas():
    df = pd.DataFrame(
        [
            {
                "Codigo": 1,
                "nombre colaborador": "Tienda A",
                "territorio": "LIMA SUR",
                "TAMAÑO": "grande",
            }
        ]
    )

    resultado, errores = preparar_dataframe(df, TDAS_COLUMNAS, "Tiendas")

    assert errores == []
    assert set(TDAS_COLUMNAS).issubset(resultado.columns)


def test_prefiere_el_primer_alias_cuando_hay_varios_presentes():
    """Si el archivo trae 'tamaño' y 'TIPOREGALO', gana el alias preferente."""
    df = pd.DataFrame(
        [
            {
                "CODIGO": 1,
                "NOMBRE_COLABORADOR": "Tienda A",
                "TERRITORIO": "LIMA SUR",
                "tamaño": "mediana",
                "TIPOREGALO": "ignorado",
            }
        ]
    )

    resultado, errores = preparar_dataframe(df, TDAS_COLUMNAS, "Tiendas")

    assert errores == []
    assert resultado.loc[0, "TipoRegalo"] == "mediana"


def test_reporta_las_columnas_faltantes_sin_lanzar_excepcion():
    df = pd.DataFrame([{"CODIGO": 1}])

    resultado, errores = preparar_dataframe(df, TDAS_COLUMNAS, "Tiendas")

    assert resultado is None
    assert errores
    assert "NombreTienda" in errores[0]


def test_el_error_de_columna_faltante_enumera_los_alias_aceptados():
    df = pd.DataFrame([{"CODIGO": 1, "NOMBRE_COLABORADOR": "T", "TERRITORIO": "LIMA"}])

    resultado, errores = preparar_dataframe(df, TDAS_COLUMNAS, "Tiendas")

    assert resultado is None
    detalle = " ".join(errores)
    assert "TipoRegalo" in detalle
    assert "TAMAÑO" in detalle


def test_rechaza_archivos_con_columnas_duplicadas():
    df = pd.DataFrame([[1, 2]], columns=["CODIGO", "CODIGO"])

    resultado, errores = preparar_dataframe(df, TDAS_COLUMNAS, "Tiendas")

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
                "tamaño": "mediana",
            }
        ]
    )

    resultado, errores = preparar_dataframe(df, TDAS_COLUMNAS, "Tiendas")

    assert errores == []
    assert not resultado.columns.duplicated().any()


# ---------------------------------------------------------------------------
# Lectura de archivos reales
# ---------------------------------------------------------------------------
def test_el_inventario_se_lee_saltando_las_filas_de_encabezado():
    archivo = construir_xlsx(
        encabezados_de(INV_COLUMNAS),
        [["01/15/2025 10:30:00 AM", "LIMA", "TIPO1", "ART-1", "Taza", 5]],
        filas_basura=2,
    )

    inv, errores = cargar_inventario(archivo)

    assert errores == []
    assert len(inv) == 1
    assert inv.loc[0, "CodigoArticulo"] == "ART-1"


def test_las_tiendas_se_leen_desde_la_primera_fila():
    archivo = construir_xlsx(
        encabezados_de(TDAS_COLUMNAS), [[1, "Tienda A", "LIMA", "TIPO1"]]
    )

    tdas, errores = cargar_tiendas(archivo)

    assert errores == []
    assert tdas.loc[0, "NombreTienda"] == "Tienda A"


def test_elige_la_hoja_con_los_encabezados_esperados():
    """La plantilla real trae primero una tabla dinámica; hay que ignorarla."""
    archivo = construir_xlsx(
        encabezados_de(TDAS_COLUMNAS),
        [[1, "Tienda A", "LIMA SUR", "mediana"]],
        hoja_previa=[
            ["lima", "LIMA SUR"],
            [],
            ["Cuenta de CODIGO", "Etiquetas de columna"],
            ["Etiquetas de fila", "grande", "mediana"],
        ],
        nombre_hoja="tiendas",
    )

    tdas, errores = cargar_tiendas(archivo)

    assert errores == []
    assert tdas.attrs["hoja"] == "tiendas"
    assert tdas.loc[0, "NombreTienda"] == "Tienda A"


def test_el_inventario_tambien_elige_la_hoja_correcta():
    archivo = construir_xlsx(
        encabezados_de(INV_COLUMNAS),
        [["01/15/2025 10:30:00 AM", "LIMA", "TIPO1", "ART-1", "Taza", 5]],
        filas_basura=2,
        hoja_previa=[["Notas del reporte"]],
        nombre_hoja="Productos",
    )

    inv, errores = cargar_inventario(archivo)

    assert errores == []
    assert inv.attrs["hoja"] == "Productos"
    assert inv.loc[0, "CodigoArticulo"] == "ART-1"


def test_flujo_completo_desde_los_excel_hasta_la_asignacion():
    from asignador_regalos import ejecutar_asignacion

    inv_file = construir_xlsx(
        encabezados_de(INV_COLUMNAS),
        [
            ["01/15/2025 10:30:00 AM", "LIMA", "TIPO1", "ART-1", "Taza roja", 3],
            ["01/16/2025 10:30:00 AM", "LIMA", "TIPO1", "ART-2", "Polo azul", 3],
        ],
        filas_basura=2,
    )
    tdas_file = construir_xlsx(
        encabezados_de(TDAS_COLUMNAS),
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
