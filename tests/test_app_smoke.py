"""Smoke tests de la interfaz: la app debe cargar y reaccionar sin excepciones."""
import sys
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

APP = str(RAIZ / "app.py")


def test_la_app_carga_sin_excepciones():
    at = AppTest.from_file(APP).run()

    assert not at.exception
    assert at.title[0].value == "Asignación automatizada de regalos a tiendas"


def test_las_opciones_de_estrategia_estan_disponibles():
    at = AppTest.from_file(APP).run()

    estrategias = at.selectbox[0].options
    assert estrategias == ["Sobrantes", "Novedades", "AltoStock", "Equitativo"]


def test_ya_no_se_elige_el_numero_de_regalos_en_la_interfaz():
    """La cantidad la define la columna 'Regalo adicional', no un selector."""
    at = AppTest.from_file(APP).run()

    assert len(at.selectbox) == 1


def test_avisa_cuando_la_plantilla_no_trae_la_columna_de_regalo_adicional():
    import pandas as pd

    import app as modulo_app

    tdas = pd.DataFrame([{"IDTienda": 1}])
    tdas.attrs["columnas_opcionales_ausentes"] = ["TipoRegaloAdicional"]

    assert modulo_app.aviso_de_columna_ausente(tdas) is True
    assert modulo_app.aviso_de_columna_ausente(pd.DataFrame([{"IDTienda": 1}])) is False


def test_pide_los_archivos_si_se_genera_sin_subirlos():
    at = AppTest.from_file(APP).run()

    at.button[0].click().run()

    assert not at.exception
    assert any("Sube los dos archivos" in e.value for e in at.error)


def test_la_barra_muestra_la_huella_y_la_version_reales():
    from info_version import VERSION, huella_codigo

    at = AppTest.from_file(APP).run()

    marcado = " ".join(m.value for m in at.markdown)
    assert huella_codigo() in marcado
    assert f"v{VERSION}" in marcado


def test_el_panel_lateral_reporta_todos_los_chequeos_en_verde():
    at = AppTest.from_file(APP).run()

    marcado = " ".join(m.value for m in at.markdown)
    assert "Conservación de stock (Equitativo)" in marcado
    assert "✗" not in marcado
    assert not at.error


def test_el_boton_de_ejemplo_ejecuta_una_asignacion_completa():
    at = AppTest.from_file(APP).run()

    boton_demo = next(b for b in at.button if "ejemplo" in b.label.lower())
    boton_demo.click().run()

    assert not at.exception
    assert any("completada con éxito" in s.value for s in at.success)
    # 5 tiendas de ejemplo, todas deben recibir al menos un regalo
    assert at.metric[1].value == "5"
    # La tienda 101 de los datos de ejemplo pide un adicional de TIPO2.
    adicionales = next(
        m for m in at.metric if m.label == "Regalos adicionales entregados"
    )
    assert adicionales.value == "1"
