# app.py
import html

import streamlit as st

from asignador_regalos import ejecutar_asignacion
from carga_datos import cargar_inventario, cargar_tiendas
from info_version import (
    VERSION,
    autochequeo,
    datos_de_ejemplo,
    hora_de_inicio_texto,
    huella_codigo,
    info_entorno,
)

st.set_page_config(page_title="Asignador de regalos", layout="wide")

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

ESTILOS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&display=swap');

.dpl-strip {
  --dpl-ink: #c9d3de;
  --dpl-dim: #6b7887;
  font-family: 'JetBrains Mono', ui-monospace, SFMono-Regular, Menlo, monospace;
  background:
    repeating-linear-gradient(0deg, rgba(255,255,255,.022) 0 1px, transparent 1px 3px),
    linear-gradient(135deg, #11151b 0%, #1a212a 100%);
  border: 1px solid #262f3a;
  border-left: 3px solid var(--dpl-accent);
  border-radius: 4px;
  padding: 13px 16px 12px;
  margin: 0 0 18px;
  color: var(--dpl-ink);
  line-height: 1.55;
}
.dpl-row { display: flex; flex-wrap: wrap; align-items: baseline; gap: 8px 14px; }
.dpl-row + .dpl-row { margin-top: 5px; }
.dpl-name {
  font-size: 11px; font-weight: 700; letter-spacing: .16em;
  text-transform: uppercase; color: var(--dpl-dim);
}
.dpl-status {
  margin-left: auto; font-size: 11px; font-weight: 700;
  letter-spacing: .12em; text-transform: uppercase; color: var(--dpl-accent);
}
.dpl-dot {
  display: inline-block; width: 7px; height: 7px; border-radius: 50%;
  background: var(--dpl-accent); margin-right: 7px; vertical-align: middle;
  box-shadow: 0 0 0 3px color-mix(in srgb, var(--dpl-accent) 22%, transparent);
}
.dpl-hash {
  font-size: 20px; font-weight: 700; letter-spacing: .04em;
  color: var(--dpl-accent);
}
.dpl-ver { font-size: 13px; font-weight: 500; color: var(--dpl-ink); }
.dpl-meta { font-size: 12px; color: var(--dpl-dim); }
.dpl-sep { color: #33404e; }
.dpl-chip {
  font-size: 11px; color: var(--dpl-dim);
  border: 1px solid #262f3a; border-radius: 3px; padding: 1px 7px;
}
.dpl-chip b { color: var(--dpl-ink); font-weight: 500; }
.dpl-hint { font-size: 11px; color: var(--dpl-dim); margin-top: 7px; }

.dpl-check {
  font-family: 'JetBrains Mono', ui-monospace, monospace;
  font-size: 12px; padding: 5px 0; border-bottom: 1px solid rgba(128,128,128,.16);
}
.dpl-check:last-child { border-bottom: none; }
.dpl-check .mark { font-weight: 700; margin-right: 6px; }
.dpl-check .ok { color: #1a9c5b; }
.dpl-check .bad { color: #d64545; }
.dpl-check .det { display: block; font-size: 10.5px; opacity: .65; margin-left: 18px; }
</style>
"""


def barra_de_version(chequeos):
    """Franja superior con la identidad exacta del código que está corriendo."""
    fallidos = [c for c in chequeos if not c["ok"]]
    if fallidos:
        acento, estado = "#ff5f56", f"{len(fallidos)} chequeo(s) fallidos"
    else:
        acento, estado = "#3ddc84", "operativo"

    entorno = info_entorno()
    chips = "".join(
        f'<span class="dpl-chip">{html.escape(k)} <b>{html.escape(v)}</b></span>'
        for k, v in entorno.items()
    )

    return f"""
<div class="dpl-strip" style="--dpl-accent:{acento}">
  <div class="dpl-row">
    <span class="dpl-name">Asignador de regalos</span>
    <span class="dpl-status"><span class="dpl-dot"></span>{html.escape(estado)}</span>
  </div>
  <div class="dpl-row">
    <span class="dpl-hash">{html.escape(huella_codigo())}</span>
    <span class="dpl-ver">v{html.escape(VERSION)}</span>
    <span class="dpl-sep">|</span>
    <span class="dpl-meta">activo desde {html.escape(hora_de_inicio_texto())}</span>
  </div>
  <div class="dpl-row">{chips}</div>
  <div class="dpl-hint">
    La huella se calcula del contenido de los archivos .py. Si tras un despliegue
    sigue siendo la misma, el c&oacute;digo nuevo no se carg&oacute;.
  </div>
</div>
"""


@st.cache_data(show_spinner=False)
def chequeos_cacheados():
    """El resultado sólo puede cambiar al reiniciar el proceso, es decir, al desplegar."""
    return autochequeo()


def panel_diagnostico(chequeos):
    st.sidebar.subheader("Estado del despliegue")
    st.sidebar.caption(
        "Cada punto ejecuta el motor real de asignación y comprueba que la "
        "corrección correspondiente esté activa en este despliegue."
    )
    filas = []
    for c in chequeos:
        clase, marca = ("ok", "✓") if c["ok"] else ("bad", "✗")
        filas.append(
            f'<div class="dpl-check">'
            f'<span class="mark {clase}">{marca}</span>{html.escape(c["nombre"])}'
            f'<span class="det">{html.escape(c["detalle"])}</span>'
            f"</div>"
        )
    st.sidebar.markdown("".join(filas), unsafe_allow_html=True)

    if any(not c["ok"] for c in chequeos):
        st.sidebar.error(
            "Hay chequeos fallidos: el despliegue no está corriendo el código esperado."
        )


def origen_leido(df):
    """Deja a la vista de qué hoja salieron los datos y con qué columnas.

    La plantilla de tiendas trae una tabla dinámica en la primera hoja, así que
    saber qué hoja se usó es lo primero que hay que confirmar ante una duda.
    """
    hoja = df.attrs.get("hoja")
    detalle = f"{len(df)} filas · {len(df.columns)} columnas"
    if hoja:
        return f"Hoja leída: «{hoja}» — {detalle}"
    return detalle


def mostrar_errores(errores):
    """Muestra el mensaje principal como error y el resto como detalle."""
    if errores:
        st.error(errores[0])
        for detalle in errores[1:]:
            st.caption(detalle)


# Lo que se mira primero al revisar una asignación: quién es la tienda, qué
# recibió y por qué no recibió más. El resto de la plantilla (latitud, códigos
# internos) empuja a NOTAS fuera de la pantalla si se deja el orden original.
COLUMNAS_DESTACADAS = [
    "IDTienda",
    "NombreTienda",
    "Zona",
    "TipoRegalo",
    "TipoRegaloAdicional",
    "REGALO_1",
    "DESC_REGALO_1",
    "REGALO_2",
    "DESC_REGALO_2",
    "NOTAS",
]


COLUMNAS_INVENTARIO = [
    "CodigoArticulo",
    "DescripcionArticulo",
    "ZonaElegible",
    "TipoRegalo",
    "CantidadDisponible",
    "UnidadesEntregadas",
]


def columnas_al_frente(df, destacadas=None):
    """Orden de columnas con las relevantes primero y el resto detrás.

    Solo reordena la vista: el DataFrame y el Excel descargado no se tocan.
    """
    destacadas = COLUMNAS_DESTACADAS if destacadas is None else destacadas
    presentes = [c for c in destacadas if c in df.columns]
    return presentes + [c for c in df.columns if c not in presentes]


def aviso_de_columna_ausente(tdas):
    """Avisa si la plantilla no traía la columna 'Regalo adicional'.

    Devuelve True si mostró el aviso, para poder probarlo sin levantar la app.
    """
    if "TipoRegaloAdicional" not in tdas.attrs.get("columnas_opcionales_ausentes", []):
        return False
    st.info(
        "La plantilla de tiendas no trae la columna **Regalo adicional**: "
        "cada tienda recibirá un solo regalo. Agrégala si necesitas asignar "
        "un segundo obsequio."
    )
    return True


def procesar(inv_file, tdas_file, estrategia):
    """Lee, valida y ejecuta la asignación. Devuelve None si la validación falla."""
    inv, errores_inv = cargar_inventario(inv_file)
    tdas, errores_tdas = cargar_tiendas(tdas_file)

    mostrar_errores(errores_inv)
    mostrar_errores(errores_tdas)
    if inv is None or tdas is None:
        return None

    aviso_de_columna_ausente(tdas)

    return inv, tdas, ejecutar_asignacion(inv, tdas, estrategia)


def mostrar_resultados(inv, tdas, resultado):
    asignaciones, inv_rest, reporte_txt, excel_bytes = resultado

    st.success("¡Asignación completada con éxito!")
    if "ADVERTENCIA" in reporte_txt:
        st.warning(
            "La ejecución terminó con advertencias. Revísalas en la pestaña 'Reporte'."
        )

    tiendas_con_regalo = int(asignaciones["REGALO_1"].ne("").sum())
    total_entregado = tiendas_con_regalo + int(asignaciones["REGALO_2"].ne("").sum())
    metricas = asignaciones.attrs.get("metricas", {})

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Tiendas procesadas", len(asignaciones))
    m2.metric("Tiendas con asignación", tiendas_con_regalo)
    m3.metric("Regalos entregados", total_entregado)
    m4.metric(
        "Regalos adicionales entregados",
        metricas.get("adicionales_entregados", 0),
    )

    tab1, tab2, tab3, tab4 = st.tabs(
        ["🎯 Asignaciones", "📦 Inventario Restante", "📋 Reporte", "📊 Vistas Previas"]
    )

    with tab1:
        st.caption(
            "Las columnas del resultado se muestran primero; el resto de la "
            "plantilla queda a la derecha. El archivo descargado conserva el "
            "orden original."
        )
        st.dataframe(
            asignaciones,
            column_order=columnas_al_frente(asignaciones),
            column_config={
                "NOTAS": st.column_config.TextColumn("NOTAS", width="large")
            },
            use_container_width=True,
        )
        st.download_button(
            "⬇️ Descargar asignacion_final.xlsx",
            data=excel_bytes,
            file_name="asignacion_final.xlsx",
            mime=MIME_XLSX,
            use_container_width=True,
        )

    with tab2:
        st.caption(
            "Stock que quedó sin repartir. `CantidadDisponible` ya tiene los "
            "descuentos de esta corrida y `UnidadesEntregadas` dice cuántas "
            "salieron de cada fila."
        )
        st.dataframe(
            inv_rest,
            column_order=columnas_al_frente(inv_rest, COLUMNAS_INVENTARIO),
            use_container_width=True,
        )

    with tab3:
        st.subheader("Regalos entregados por artículo y zona")
        matriz = asignaciones.attrs.get("matriz_zonas")
        if matriz is None or matriz.empty:
            st.info("No se entregó ningún regalo, así que no hay nada que resumir.")
        else:
            st.caption(
                "Es la tabla dinámica que se armaba a mano: total por artículo "
                "a la derecha y total por zona en la última fila. Va también "
                "como hoja «RegalosPorZona» del Excel descargado."
            )
            st.dataframe(matriz, use_container_width=True, hide_index=True)

        st.subheader("Resumen de la ejecución")
        # `st.code` respeta la fuente monoespaciada: el reporte alinea sus
        # cifras en columna y con la tipografía normal la alineación se rompe.
        st.code(reporte_txt, language=None)
        st.download_button(
            "⬇️ Descargar reporte.txt",
            data=reporte_txt.encode("utf-8"),
            file_name="reporte.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with tab4:
        st.info(
            "Estas son las **entradas tal como se leyeron**, antes de asignar: "
            "las cantidades son las originales, sin descuentos. El stock ya "
            "descontado está en la pestaña «Inventario Restante»."
        )
        st.subheader("Inventario leído (columnas ya renombradas)")
        st.caption(origen_leido(inv))
        st.dataframe(
            inv.head(),
            column_order=columnas_al_frente(inv, COLUMNAS_INVENTARIO),
            use_container_width=True,
        )
        st.subheader("Tiendas leídas (columnas ya renombradas)")
        st.caption(origen_leido(tdas))
        st.dataframe(
            tdas.head(),
            column_order=columnas_al_frente(tdas),
            use_container_width=True,
        )


# --- Interfaz ---
st.markdown(ESTILOS, unsafe_allow_html=True)
st.title("Asignación automatizada de regalos a tiendas")

chequeos = chequeos_cacheados()
st.markdown(barra_de_version(chequeos), unsafe_allow_html=True)
panel_diagnostico(chequeos)

col1, col2 = st.columns(2)

with col1:
    st.info("Carga los archivos de inventario y tiendas para comenzar.")
    inv_file = st.file_uploader("1. Inventario (inventario.xlsx)", type=["xlsx"])
    tdas_file = st.file_uploader("2. Tiendas (tiendas.xlsx)", type=["xlsx"])

with col2:
    st.info("Configura los parámetros de la asignación.")
    estrategia = st.selectbox(
        "Estrategia de asignación",
        ["Sobrantes", "Novedades", "AltoStock", "Equitativo"],
        help="Define qué artículos se usarán primero.",
    )
    st.caption(
        "La cantidad de regalos ya no se elige aquí: la define la columna "
        "**Regalo adicional** de la plantilla de tiendas. Si está vacía, la "
        "tienda recibe un regalo; si trae un tipo, recibe además uno de ese tipo."
    )

col_generar, col_demo = st.columns([3, 1])
generar = col_generar.button(
    "🚀 Generar Asignación", type="primary", use_container_width=True
)
demo = col_demo.button(
    "🧪 Datos de ejemplo",
    use_container_width=True,
    help="Ejecuta la asignación con datos integrados, sin subir archivos.",
)

# --- Lógica principal ---
resultado = None

if generar:
    if not inv_file or not tdas_file:
        st.error("⚠️ Sube los dos archivos requeridos antes de continuar.")
    else:
        with st.spinner("Procesando archivos y realizando asignaciones..."):
            try:
                resultado = procesar(inv_file, tdas_file, estrategia)
            except Exception as e:  # noqa: BLE001 - se muestra la traza al usuario
                st.error("Ocurrió un error inesperado durante el proceso.")
                st.exception(e)

elif demo:
    st.info(
        "Ejecución de prueba con datos integrados (5 artículos, 5 tiendas, 2 zonas). "
        "No se usó ningún archivo subido."
    )
    with st.spinner("Ejecutando con datos de ejemplo..."):
        try:
            inv_demo, tdas_demo = datos_de_ejemplo()
            resultado = (
                inv_demo,
                tdas_demo,
                ejecutar_asignacion(inv_demo, tdas_demo, estrategia),
            )
        except Exception as e:  # noqa: BLE001 - se muestra la traza al usuario
            st.error("Ocurrió un error inesperado con los datos de ejemplo.")
            st.exception(e)

if resultado is not None:
    mostrar_resultados(*resultado)
