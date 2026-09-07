"""Identidad del despliegue y autodiagnóstico.

Permite confirmar de un vistazo, desde la app ya desplegada, si el código que
está corriendo es el que se acaba de subir. La huella se calcula del contenido
real de los archivos fuente: no depende de recordar subir el número de versión.
"""
import hashlib
import platform
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import pandas as pd

from asignador_regalos import ejecutar_asignacion

VERSION = "2.0.0"

RAIZ = Path(__file__).resolve().parent

# Streamlit Cloud corre en UTC; se muestra la hora local del equipo.
ZONA_HORARIA = "America/Lima"


def _zona():
    try:
        return ZoneInfo(ZONA_HORARIA)
    except (ZoneInfoNotFoundError, KeyError):
        return None


# Se evalúa una sola vez al importar el módulo: marca el arranque del proceso,
# es decir, el momento en que el despliegue quedó activo.
_INICIO = datetime.now(_zona())


def hora_de_inicio():
    """Momento en que arrancó este proceso (equivale a la hora del despliegue)."""
    return _INICIO


def hora_de_inicio_texto():
    """Hora de arranque formateada, con la zona horaria explícita."""
    etiqueta = ZONA_HORARIA if _zona() else "hora local del servidor"
    return f"{_INICIO.strftime('%d/%m/%Y %H:%M')} ({etiqueta})"

# Archivos cuyo contenido define la huella del despliegue.
ARCHIVOS_FUENTE = [
    RAIZ / "app.py",
    RAIZ / "asignador_regalos.py",
    RAIZ / "carga_datos.py",
    RAIZ / "info_version.py",
]


def huella_codigo(archivos=None):
    """Hash corto del contenido de los archivos fuente.

    Cambia siempre que cambia el código, así que sirve para verificar que un
    despliegue realmente tomó la versión nueva. Los archivos ausentes se
    incorporan al hash como tales en vez de provocar un error.
    """
    digest = hashlib.sha256()
    for ruta in sorted(archivos or ARCHIVOS_FUENTE, key=lambda p: Path(p).name):
        ruta = Path(ruta)
        digest.update(ruta.name.encode("utf-8"))
        try:
            digest.update(ruta.read_bytes())
        except OSError:
            digest.update(b"<ausente>")
    return digest.hexdigest()[:7]


def info_entorno():
    """Versiones instaladas de las dependencias críticas."""
    import openpyxl
    import streamlit

    return {
        "python": platform.python_version(),
        "streamlit": streamlit.__version__,
        "pandas": pd.__version__,
        "openpyxl": openpyxl.__version__,
    }


def datos_de_ejemplo():
    """Inventario y tiendas de prueba para validar la app sin subir archivos."""
    inv = pd.DataFrame(
        [
            ("LIMA", "TIPO1", "ART-101", "Taza cerámica", "01/15/2025 09:00:00 AM", 4),
            ("LIMA", "TIPO1", "ART-102", "Polo algodón", "01/20/2025 09:00:00 AM", 3),
            ("LIMA", "TIPO2", "ART-201", "Mochila", "01/18/2025 09:00:00 AM", 2),
            ("NORTE", "TIPO1", "ART-301", "Termo acero", "01/12/2025 09:00:00 AM", 5),
            ("NORTE", "TIPO1", "ART-302", "Gorra", "01/22/2025 09:00:00 AM", 5),
        ],
        columns=[
            "ZonaElegible",
            "TipoRegalo",
            "CodigoArticulo",
            "DescripcionArticulo",
            "FechaIngreso",
            "CantidadDisponible",
        ],
    )
    tdas = pd.DataFrame(
        [
            (101, "Tienda Centro", "LIMA", "TIPO1"),
            (102, "Tienda Miraflores", "LIMA", "TIPO1"),
            (103, "Tienda Surco", "LIMA", "TIPO2"),
            (201, "Tienda Trujillo", "NORTE", "TIPO1"),
            (202, "Tienda Chiclayo", "NORTE", "TIPO1"),
        ],
        columns=["IDTienda", "NombreTienda", "Zona", "TipoRegalo"],
    )
    return inv, tdas


# ---------------------------------------------------------------------------
# Autochequeo: cada función ejecuta el motor real y devuelve un detalle
# ---------------------------------------------------------------------------
def _inventario(filas):
    return pd.DataFrame(
        filas,
        columns=[
            "ZonaElegible",
            "TipoRegalo",
            "CodigoArticulo",
            "DescripcionArticulo",
            "FechaIngreso",
            "CantidadDisponible",
        ],
    )


def _tiendas(filas):
    return pd.DataFrame(
        filas, columns=["IDTienda", "NombreTienda", "Zona", "TipoRegalo"]
    )


def _chequeo_stock_equitativo():
    """El bug crítico: 'Equitativo' descontaba stock a las filas equivocadas."""
    inv = _inventario(
        [
            ("A", "T1", "A1", "d", "01/15/2025 09:00:00 AM", 10),
            ("A", "T1", "A2", "d", "01/15/2025 09:00:00 AM", 10),
            ("B", "T1", "B1", "d", "01/15/2025 09:00:00 AM", 100),
            ("B", "T1", "B2", "d", "01/15/2025 09:00:00 AM", 100),
        ]
    )
    tdas = _tiendas([(1, "TA", "A", "T1"), (2, "TB", "B", "T1")])

    asignaciones, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, 1, "Equitativo")
    entregado = int(asignaciones["REGALO_1"].ne("").sum())
    restante = int(inv_rest["CantidadDisponible"].sum())

    ok = (entregado + restante) == 220
    return ok, f"220 unidades iniciales = {entregado} entregadas + {restante} restantes"


def _chequeo_columnas_descripcion():
    inv = _inventario([("A", "T1", "A1", "Taza roja", "01/15/2025 09:00:00 AM", 5)])
    tdas = _tiendas([(1, "TA", "A", "T1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")
    ok = "DESC_REGALO_1" in asignaciones.columns
    presentes = [c for c in asignaciones.columns if c.startswith("DESC_")]
    return ok, f"Columnas de descripción: {presentes or 'ninguna'}"


def _chequeo_fechas_flexibles():
    inv = _inventario([("A", "T1", "A1", "d", "2025-08-28 00:00:00", 5)])
    tdas = _tiendas([(1, "TA", "A", "T1")])

    _, inv_rest, _, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")
    ok = bool(inv_rest["FechaIngreso"].notna().all())
    return ok, "Fecha ISO '2025-08-28' interpretada correctamente"


def _chequeo_cruce_mayusculas():
    inv = _inventario([("  lima ", "tipo1", "A1", "d", "01/15/2025 09:00:00 AM", 5)])
    tdas = _tiendas([(1, "TA", "LIMA", "TIPO1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 1, "Sobrantes")
    ok = asignaciones.loc[0, "REGALO_1"] == "A1"
    return ok, "Zona '  lima ' cruza con 'LIMA'"


def _chequeo_regalos_distintos():
    inv = _inventario(
        [
            ("A", "T1", "A1", "Taza", "01/15/2025 09:00:00 AM", 5),
            ("A", "T1", "A2", "Polo", "01/15/2025 09:00:00 AM", 5),
        ]
    )
    tdas = _tiendas([(1, "TA", "A", "T1")])

    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 2, "Sobrantes")
    r1, r2 = asignaciones.loc[0, "REGALO_1"], asignaciones.loc[0, "REGALO_2"]
    return r1 != r2, f"Con 2 regalos entrega '{r1}' y '{r2}'"


CHEQUEOS = [
    ("Conservación de stock (Equitativo)", _chequeo_stock_equitativo),
    ("Columnas de descripción", _chequeo_columnas_descripcion),
    ("Parseo flexible de fechas", _chequeo_fechas_flexibles),
    ("Cruce insensible a mayúsculas", _chequeo_cruce_mayusculas),
    ("Dos regalos distintos", _chequeo_regalos_distintos),
]


def autochequeo():
    """Ejecuta el motor real y verifica que las correcciones estén activas.

    Nunca lanza: un chequeo roto se reporta como fallo para que la app siga
    siendo utilizable y el problema quede visible.
    """
    resultados = []
    for nombre, funcion in CHEQUEOS:
        try:
            ok, detalle = funcion()
        except Exception as e:  # noqa: BLE001 - el detalle se muestra al usuario
            ok, detalle = False, f"Error al ejecutar el chequeo: {e}"
        resultados.append({"nombre": nombre, "ok": bool(ok), "detalle": detalle})
    return resultados
