"""Pruebas del panel de versión y autodiagnóstico del despliegue."""
import re
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from info_version import (
    VERSION,
    autochequeo,
    datos_de_ejemplo,
    hora_de_inicio,
    hora_de_inicio_texto,
    huella_codigo,
    info_entorno,
)


# ---------------------------------------------------------------------------
# Hora de arranque del proceso (= momento del despliegue)
# ---------------------------------------------------------------------------
def test_la_hora_de_inicio_es_la_misma_en_todo_el_proceso():
    """Debe reflejar el arranque del contenedor, no el instante de la consulta."""
    assert hora_de_inicio() is hora_de_inicio()


def test_la_hora_de_inicio_se_muestra_con_su_zona_horaria():
    texto = hora_de_inicio_texto()

    assert re.match(r"\d{2}/\d{2}/\d{4} \d{2}:\d{2}", texto)
    assert texto.strip().endswith(")")


# ---------------------------------------------------------------------------
# Huella del código
# ---------------------------------------------------------------------------
def test_la_huella_es_un_hash_corto_hexadecimal():
    assert re.fullmatch(r"[0-9a-f]{7}", huella_codigo())


def test_la_huella_es_estable_entre_llamadas():
    assert huella_codigo() == huella_codigo()


def test_la_huella_cambia_si_cambia_el_contenido(tmp_path):
    archivo = tmp_path / "modulo.py"

    archivo.write_text("valor = 1", encoding="utf-8")
    antes = huella_codigo([archivo])

    archivo.write_text("valor = 2", encoding="utf-8")
    despues = huella_codigo([archivo])

    assert antes != despues


def test_la_huella_ignora_archivos_inexistentes(tmp_path):
    """Un despliegue parcial no debe romper la app, solo cambiar la huella."""
    assert re.fullmatch(r"[0-9a-f]{7}", huella_codigo([tmp_path / "no_existe.py"]))


def test_la_version_sigue_el_formato_semantico():
    assert re.fullmatch(r"\d+\.\d+\.\d+", VERSION)


# ---------------------------------------------------------------------------
# Entorno
# ---------------------------------------------------------------------------
def test_el_entorno_reporta_las_librerias_criticas():
    entorno = info_entorno()

    for libreria in ("streamlit", "pandas", "openpyxl", "python"):
        assert libreria in entorno
        assert entorno[libreria]


# ---------------------------------------------------------------------------
# Autochequeo
# ---------------------------------------------------------------------------
def test_el_autochequeo_confirma_todas_las_correcciones():
    resultados = autochequeo()

    assert len(resultados) >= 4
    fallidos = [r["nombre"] for r in resultados if not r["ok"]]
    assert fallidos == []


def test_cada_chequeo_trae_nombre_y_detalle():
    for resultado in autochequeo():
        assert resultado["nombre"]
        assert resultado["detalle"]
        assert isinstance(resultado["ok"], bool)


def test_el_autochequeo_no_lanza_excepciones_aunque_algo_falle(monkeypatch):
    """Un chequeo roto debe reportarse como fallo, no tumbar la app."""
    import info_version

    def explota(*args, **kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr(info_version, "ejecutar_asignacion", explota)

    resultados = info_version.autochequeo()

    assert any(not r["ok"] for r in resultados)
    assert any("boom" in r["detalle"] for r in resultados)


# ---------------------------------------------------------------------------
# Datos de ejemplo
# ---------------------------------------------------------------------------
def test_los_datos_de_ejemplo_producen_una_asignacion_completa():
    from asignador_regalos import ejecutar_asignacion

    inv, tdas = datos_de_ejemplo()
    asignaciones, _, _, _ = ejecutar_asignacion(inv, tdas, 2, "Equitativo")

    assert len(asignaciones) > 0
    assert asignaciones["REGALO_1"].ne("").all()
