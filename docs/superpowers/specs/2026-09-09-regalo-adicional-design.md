# Regalo adicional por tienda

**Fecha:** 2026-09-09
**Estado:** aprobado, pendiente de plan de implementación

---

## Problema

Hoy la cantidad de regalos es un parámetro global de la interfaz: un selectbox
"N° de regalos por tienda" con las opciones 1 y 2, que se aplica por igual a
las 1840 tiendas. No hay forma de decir "estas 200 tiendas reciben dos regalos
y el resto uno", ni de pedir que el segundo regalo salga de un segmento
distinto al de la tienda.

La plantilla de tiendas ya resuelve el problema desde el lado de los datos: la
hoja `tiendas` de `TIENDAS_PLANTILLA_FINAL (2).xlsx` trae una columna nueva
llamada **`Regalo adicional`**, al lado de `tamaño`, que contiene un tipo de
regalo (`pequeña`, `mediana`, `grande`, …) con el mismo vocabulario que
`tamaño`. Está vacía en las 1840 filas porque todavía no se usa.

## Objetivo

Que la cantidad de regalos de cada tienda salga de sus datos y no de un control
global:

- Celda `Regalo adicional` vacía → la tienda recibe **1 regalo**, del tipo que
  indica su columna `tamaño`.
- Celda con un tipo → la tienda recibe **2 regalos**: el primero del pozo de
  inventario de su `tamaño`, el segundo del pozo del tipo indicado en
  `Regalo adicional`. Los dos tipos pueden ser distintos.

El selectbox de 1 ó 2 regalos desaparece de la interfaz y el parámetro
`n_regalos` desaparece del motor.

## Decisiones tomadas

| Pregunta | Decisión |
| --- | --- |
| ¿El valor de `Regalo adicional` define el pozo del 2° regalo, o solo marca "dale 2"? | **Define su propio pozo.** Una tienda `pequeña` con adicional `mediana` recibe un artículo del stock PEQUEÑO y otro del stock MEDIANO, ambos de su zona. |
| ¿Qué pasa si solo se consigue uno de los dos? | **Se entrega el que sí hay, con nota.** Cuenta como tienda asignada y suma al conteo de parciales del reporte. |
| ¿Qué pasa si el archivo no trae la columna? | **Se procesa igual**, asumiendo la columna vacía (todas las tiendas con 1 regalo). La app avisa en pantalla que no la encontró. La columna es opcional, no obligatoria. |

## Fuera de alcance

El desajuste de vocabulario entre plantillas (`pequeña` en tiendas contra
`PEQUEÑO` en inventario) no se toca. El cruce sigue siendo exacto salvo
mayúsculas, espacios y acentos, y corregir los datos de origen queda del lado
del usuario, que trabajará con la plantilla ya normalizada.

---

## Arquitectura

La restricción de diseño es doble. Primero, el motor deja de razonar en
"cuántos regalos" y pasa a razonar en "qué tipo pide esta tienda en cada
ronda": el primer regalo sale del pozo de su `tamaño` y el adicional del pozo
que nombra su celda `Regalo adicional`.

Segundo, y esto define la forma del motor: **los primeros regalos tienen
prioridad absoluta sobre los adicionales**. Por eso la asignación se organiza
en dos pasadas sucesivas sobre las tiendas de cada zona, en vez de servirle a
cada tienda todo lo que pide de una sola vez. La unidad de trabajo pasa a ser
"entregar un regalo de un tipo", y las dos pasadas la invocan con distinto
pozo. Con eso hay un solo camino de código en lugar de dos ramas
(`n == 1` / `n == 2`), y una eventual tercera ronda de regalos sería una
pasada más, sin cambiar la estructura.

Los cuatro módulos conservan sus responsabilidades actuales:
`carga_datos.py` mapea y valida sin saber de Streamlit, `asignador_regalos.py`
asigna sin saber de Excel de entrada, `app.py` solo presenta, e
`info_version.py` verifica el despliegue ejecutando el motor real.

### 1. Carga — `carga_datos.py`

Se agrega el concepto de **columna opcional**, que hoy no existe:

```python
TDAS_COLUMNAS_OPCIONALES = {
    "TipoRegaloAdicional": ("REGALO ADICIONAL", "TIPO REGALO ADICIONAL"),
}
```

`normalizar_encabezado` ya ignora mayúsculas, acentos, espacios y guiones
bajos, así que `Regalo adicional`, `REGALO_ADICIONAL` y `regaloadicional`
calzan con el primer alias sin necesidad de listarlos.

`preparar_dataframe` recibe un parámetro nuevo `opcionales` (por defecto
vacío) y lo trata así:

- **Presente:** se renombra al nombre interno, igual que una columna obligatoria.
- **Ausente:** *no* es un error. Se crea la columna con cadena vacía y el
  nombre interno se anota en `df.attrs["columnas_opcionales_ausentes"]`.

El motor entonces siempre encuentra `TipoRegaloAdicional`, y solo un módulo
—el de carga— decide qué hacer cuando falta. La contrapartida es que, con una
plantilla vieja, el Excel de salida gana una columna vacía llamada
`TipoRegaloAdicional`; es un costo aceptable y además deja a la vista que la
plantilla está desactualizada.

`cargar_tiendas` pasa `TDAS_COLUMNAS_OPCIONALES`. `cargar_inventario` no
cambia. `elegir_hoja` tiene en cuenta también los alias opcionales al puntuar
las hojas, de modo que la columna nueva ayude a identificar la hoja correcta
en vez de ser ignorada.

### 2. Motor — `asignador_regalos.py`

**Firma:** `ejecutar_asignacion(inv, tdas, estrategia)`. Se elimina
`n_regalos`. Es un cambio incompatible y obliga a actualizar `app.py`,
`info_version.py` y las pruebas.

**Clave de cruce nueva:** `COL_TIPO_ADIC_KEY = "_TipoAdicKey"`, generada con
la misma `clave_normalizada` que las demás (sin acentos, sin espacios,
mayúsculas) a partir de `TipoRegaloAdicional`. Se elimina junto a las otras
auxiliares antes de exportar.

**Dos pasadas.** El primer regalo de **todas** las tiendas tiene prioridad
sobre cualquier regalo adicional. Dentro de cada zona se recorre a las tiendas
dos veces:

- **Pasada 1 — primeros regalos.** Cada tienda toma una unidad del pozo de su
  `tamaño`. Ninguna tienda toca todavía el pozo del adicional.
- **Pasada 2 — regalos adicionales.** Solo las tiendas con la celda
  `Regalo adicional` no vacía toman una unidad del pozo de ese tipo, del stock
  que haya sobrado de la pasada 1.

Así, si el stock de un tipo no alcanza, lo que se pierde son regalos
adicionales y nunca el primer regalo de otra tienda. Como los pozos de
inventario están acotados por zona y ninguna zona compite con otra por stock,
hacer las dos pasadas **dentro de cada zona** da exactamente el mismo
resultado que hacerlas globalmente sobre todas las zonas; se hace por zona
porque conserva la estructura del bucle actual.

Una tienda cuyo primer regalo falló en la pasada 1 **igual participa** en la
pasada 2: llegado ese punto ya no hay primeros regalos en riesgo, así que
negarle el adicional solo dejaría stock sin repartir.

**Función de asignación.** `intentar_asignar_para_tienda(inv_tipo, n_regalos)`
—que servía los dos regalos de una vez— ya no calza con el esquema de dos
pasadas. Se reemplaza por una función que entrega **un** regalo:

```python
tomar_regalo(inv_tipo, codigos_excluidos=())
    → (ok, codigo, descripcion, inv_tipo_actualizado)
```

Toma una unidad del pozo recibido, prefiriendo un artículo que no esté en
`codigos_excluidos`. No muta el DataFrame recibido. La pasada 1 la llama sin
exclusiones; la pasada 2 la llama pasándole el código que la tienda ya
recibió.

Ese parámetro es lo que conserva la **variedad** cuando el tipo adicional es
igual al de la tienda (una `pequeña` con adicional `pequeña`), con el mismo
orden de preferencia que hoy:

1. Un artículo distinto del ya entregado.
2. Si no hay otro artículo con stock, se repite el mismo.

**Llenado de ranuras.** Los regalos conseguidos ocupan `REGALO_1` y `REGALO_2`
en el orden en que se obtuvieron, sin dejar huecos. Por lo tanto `REGALO_1`
nunca queda vacío si hubo algo que entregar, incluso cuando lo que faltó fue
el regalo principal. Esto mantiene válida la métrica "tiendas con asignación",
que se calcula como `REGALO_1 != ""`.

**Notas y excepciones:**

| Caso | `NOTAS` | ¿Excepción? |
| --- | --- | --- |
| Los 2 entregados | *(vacío)* | No |
| Falta el adicional | `Asignación parcial: sin stock del regalo adicional 'mediana' en la zona LIMA SUR` | No, cuenta como parcial |
| Falta el principal, hay adicional | `Asignación parcial: sin stock del tipo 'pequeña' en la zona LIMA SUR` | No, cuenta como parcial |
| No se consiguió nada | Motivo nombrando el o los tipos que faltaron y la zona | Sí |
| Zona sin ningún inventario | `No hay inventario disponible en la zona X` (igual que hoy) | Sí |

**Orden de servicio.** Se recorre zona por zona y, dentro de cada zona, se
hacen las dos pasadas descritas arriba, cada una recorriendo las tiendas en el
orden del archivo. No hay reserva previa ni optimización global: dentro de una
misma pasada, si el stock de un tipo se agota, las tiendas que quedan al final
del archivo son las que se quedan sin ese regalo.

Esto **cambia respecto de hoy**. Con el `n_regalos = 2` actual, una tienda del
comienzo del archivo se lleva dos unidades mientras una del final puede
quedarse sin ninguna. Con las dos pasadas, esa segunda unidad solo se entrega
una vez que todas las tiendas de la zona tuvieron su oportunidad de recibir la
primera.

**Reporte.** La línea `NumeroRegalosPorTienda: {n}` ya no tiene sentido y se
reemplaza por dos (los números son ilustrativos):

```
Tiendas que piden regalo adicional: 214
Regalos adicionales entregados: 198
```

El resto del reporte (estrategia, tiendas procesadas, tiendas con asignación,
parciales, total de regalos, unidades restantes, advertencias de fechas,
detalle de excepciones) se mantiene sin cambios.

**Contadores para la UI.** "Regalos adicionales entregados" **no** se puede
derivar de `REGALO_2 != ""`: cuando falta el principal, el regalo adicional
ocupa `REGALO_1` y ese conteo lo perdería. El motor, que sí sabe qué tipo
sirvió en cada caso, publica los contadores en
`df_tiendas_final.attrs["metricas"]`:

```python
{"piden_adicional": int, "adicionales_entregados": int, "parciales": int}
```

La UI y el reporte leen de ahí. Es el mismo mecanismo que ya se usa para
`attrs["hoja"]`.

### 3. Interfaz — `app.py`

- Se elimina el selectbox `n_regalos`. La columna de parámetros queda solo con
  la estrategia, más una nota: la cantidad de regalos la define la columna
  `Regalo adicional` de la plantilla de tiendas.
- `procesar(inv_file, tdas_file, estrategia)` y la rama de datos de ejemplo se
  ajustan a la firma nueva.
- Si `df.attrs["columnas_opcionales_ausentes"]` trae `TipoRegaloAdicional`, se
  muestra un `st.info` avisando que no se encontró la columna y que por eso
  cada tienda recibe un solo regalo.
- Se agrega una cuarta métrica, "Regalos adicionales entregados", leída de
  `asignaciones.attrs["metricas"]`, junto a las tres actuales.

### 4. Autodiagnóstico — `info_version.py`

- Los helpers `_tiendas()` y `datos_de_ejemplo()` incorporan la columna
  `TipoRegaloAdicional`. Los datos de ejemplo incluyen al menos una tienda con
  regalo adicional de un tipo distinto al suyo, para que el botón "Datos de
  ejemplo" ejercite el caso nuevo.
- El chequeo `_chequeo_regalos_distintos` ("Dos regalos distintos") pasa a ser
  **"Regalo adicional de otro tipo"**: una tienda `T1` con adicional `T2` debe
  recibir un artículo de cada pozo, y el chequeo verifica que los dos códigos
  correspondan a artículos de tipos distintos.
- Se agrega el chequeo **"Columna opcional ausente"**: un DataFrame de tiendas
  sin `TipoRegaloAdicional` debe procesarse sin error, con todas las tiendas en
  un regalo.
- Se agrega el chequeo **"Prioridad del primer regalo"**, que reproduce el caso
  de las dos tiendas y dos unidades descrito en las pruebas. Es el que confirma
  desde el despliegue que la regla de prioridad está activa.
- `VERSION` sube a `3.0.0`: cambia el contrato del archivo de entrada y la
  firma pública del motor.

### 5. Pruebas — `tests/`

Se actualizan todas las llamadas existentes a `ejecutar_asignacion` para la
firma sin `n_regalos`, y se agregan casos para:

- **Prioridad del primer regalo** (la prueba central del cambio): misma zona,
  un tipo `T1` con **2 unidades** de stock y dos tiendas de tipo `T1`, donde la
  **primera** del archivo pide además un adicional de `T1`. El resultado
  esperado es una unidad para cada tienda: la primera con su regalo y una nota
  de parcial por el adicional, la segunda con su regalo. La lógica anterior le
  daba las dos unidades a la primera tienda y dejaba a la segunda sin nada.
- **Prioridad entre zonas independientes:** el resultado de hacer las dos
  pasadas por zona coincide con hacerlas globalmente.
- Regalo adicional de un tipo **distinto**: un artículo de cada pozo.
- Regalo adicional del **mismo** tipo: dos artículos distintos (variedad).
- Tienda **sin primer regalo pero con adicional disponible**: participa en la
  pasada 2 y recibe el adicional en `REGALO_1`.
- Parcial por falta del **adicional**: `REGALO_1` lleno, `REGALO_2` vacío, nota.
- Parcial por falta del **principal**: `REGALO_1` lleno con el adicional, nota.
- Celda **vacía** y celda con **solo espacios**: un solo regalo, sin nota.
- Tipo adicional **inexistente** en la zona: parcial con nota, no excepción.
- **Columna ausente** en el Excel de tiendas: carga sin error, columna creada
  vacía y el faltante anotado en `attrs`.
- Conservación de stock: unidades iniciales = entregadas + restantes, con
  regalos adicionales en juego.

### 6. Documentación — `README.md`

- Tabla de `tiendas.xlsx`: se agrega la fila `TipoRegaloAdicional` /
  `REGALO ADICIONAL`, marcada como opcional.
- La sección "Regalos por tienda" se reescribe: la cantidad ya no se elige en
  la interfaz, sale de la columna. Su diagrama se reemplaza por uno que
  refleje las **dos pasadas**, los dos pozos independientes y los tres
  desenlaces (dos regalos, parcial, sin asignación).
- Se documenta explícitamente la regla de prioridad: nadie recibe un segundo
  regalo mientras queden tiendas de su zona sin el primero.
- Se corrige la sección "Cómo asigna" para mencionar el segundo pozo.
- Se quita la mención al selector de 1 ó 2 regalos.

---

## Flujo de datos

```mermaid
flowchart TD
    A["tiendas.xlsx<br/>tamaño + Regalo adicional"] --> B["cargar_tiendas()<br/>columna opcional: mapea o crea vacía"]
    C["inventario.xlsx"] --> D["cargar_inventario()"]
    B --> E["ejecutar_asignacion(inv, tdas, estrategia)"]
    D --> E
    E --> F["por zona: pozos de inventario por tipo"]
    F --> G["PASADA 1: toda tienda toma 1 unidad<br/>del pozo de su tamaño"]
    G --> H["PASADA 2: las tiendas con Regalo adicional<br/>toman del pozo de ese tipo, con lo que sobró"]
    H --> I["REGALO_1 / REGALO_2 en orden de obtención"]
    H --> J["tipos sin stock → NOTAS o excepción"]
    I --> K["asignacion_final.xlsx + reporte.txt"]
    J --> K
```

## Manejo de errores

| Situación | Comportamiento |
| --- | --- |
| Columna `Regalo adicional` ausente | Se crea vacía, aviso en la UI, proceso continúa |
| Celda vacía o con espacios | Un solo regalo, sin nota |
| Tipo adicional que no existe en el inventario de la zona | Parcial con nota; la tienda cuenta como asignada |
| Tipo adicional agotado durante la corrida | Idéntico al caso anterior |
| Ninguno de los dos tipos disponible | Excepción, `REGALO_1` y `REGALO_2` vacíos, motivo en `NOTAS` y en el reporte |
| Columnas obligatorias faltantes | Sin cambios: la carga falla con el detalle de lo que falta |
