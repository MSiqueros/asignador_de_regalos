import pandas as pd
import io
from datetime import datetime
from collections import deque
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font

# ============================
# Funciones auxiliares
# ============================
def normalizar_texto(df):
    """Limpia los espacios en blanco de todas las columnas de texto."""
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def ordenar_por_estrategia(df_inv, estrategia):
    """Ordena el DataFrame de inventario según la estrategia seleccionada."""
    if estrategia == "Sobrantes":
        return df_inv.sort_values(by=["TipoRegalo", "FechaIngreso", "CantidadDisponible"], ascending=[True, True, True])
    elif estrategia == "Novedades":
        return df_inv.sort_values(by=["TipoRegalo", "FechaIngreso", "CantidadDisponible"], ascending=[True, False, False])
    elif estrategia == "AltoStock":
        return df_inv.sort_values(by=["TipoRegalo", "CantidadDisponible", "FechaIngreso"], ascending=[True, False, True])
    elif estrategia == "Equitativo":
        base = df_inv.sort_values(by=["TipoRegalo", "CodigoArticulo"]).copy()
        rows = []
        for _, g in base.groupby("TipoRegalo", sort=False):
            dq = deque(g.to_dict("records"))
            while dq:
                rows.append(dq.popleft())
                if dq:
                    dq.rotate(-1)
        return pd.DataFrame(rows)
    else:
        return df_inv

def intentar_asignar_para_tienda(df_inv_tipo, n_regalos):
    """
    Intenta asignar n_regalos de un inventario de un tipo específico.
    Devuelve (éxito, códigos, descripciones, inventario_actualizado).
    """
    inv = df_inv_tipo.copy()

    def tomar(idx, unidades):
        inv.loc[idx, "CantidadDisponible"] -= unidades

    if n_regalos == 1:
        candidatos = inv[inv["CantidadDisponible"] >= 1]
        if not candidatos.empty:
            idx = candidatos.index[0]
            row = inv.loc[idx]
            tomar(idx, 1)
            return True, [row["CodigoArticulo"]], [row["DescripcionArticulo"]], inv
        return False, [], [], df_inv_tipo

    if n_regalos == 2:
        cand2 = inv[inv["CantidadDisponible"] >= 2]
        if not cand2.empty:
            idx = cand2.index[0]
            row = inv.loc[idx]
            tomar(idx, 2)
            return True, [row["CodigoArticulo"], row["CodigoArticulo"]], [row["DescripcionArticulo"], row["DescripcionArticulo"]], inv

        cand1 = inv[inv["CantidadDisponible"] >= 1]
        if len(cand1) >= 2:
            idx1, idx2 = cand1.index[0], cand1.index[1]
            r1, r2 = inv.loc[idx1], inv.loc[idx2]
            tomar(idx1, 1); tomar(idx2, 1)
            return True, [r1["CodigoArticulo"], r2["CodigoArticulo"]], [r1["DescripcionArticulo"], r2["DescripcionArticulo"]], inv

    return False, [], [], df_inv_tipo

# ============================
# Función principal
# ============================
def ejecutar_asignacion(inv, tdas, n_regalos, estrategia):
    """
    Función principal que orquesta la asignación con lógica de fallback.
    """
    # 1. Preparación de datos
    inv["FechaIngreso"] = pd.to_datetime(inv["FechaIngreso"], errors="coerce", format="%m/%d/%Y %I:%M:%S %p")
    inv = normalizar_texto(inv)
    tdas = normalizar_texto(tdas)
    inv["CantidadDisponible"] = pd.to_numeric(inv["CantidadDisponible"], errors="coerce").fillna(0).astype('Int64')
    inv = inv[inv["CantidadDisponible"] > 0].copy()

    # 2. Inicializar las columnas de asignación en la plantilla de tiendas
    # Se inicializan las columnas asegurando que su tipo de dato sea siempre texto (str)
    tdas["REGALO_1"] = pd.Series("", index=tdas.index, dtype=str)
    tdas["REGALO_2"] = pd.Series("", index=tdas.index, dtype=str)
    tdas["NOTAS"] = pd.Series("", index=tdas.index, dtype=str)

    excepciones = []
    inv_actualizado = inv.copy()

    # 3. Iterar por cada zona
    zonas = sorted(list(tdas["Zona"].unique()))
    for zona in zonas:
        tiendas_z = tdas[tdas["Zona"] == zona].copy()
        inv_z = inv_actualizado[inv_actualizado["ZonaElegible"] == zona].copy()

        if inv_z.empty:
            for _, rowt in tiendas_z.iterrows():
                excepciones.append({"IDTienda": rowt["IDTienda"], "NombreTienda": rowt["NombreTienda"], "Zona": zona, "Motivo": f"No hay inventario disponible en la zona {zona}"})
            continue

        inv_z_ord = ordenar_por_estrategia(inv_z, estrategia)
        inv_por_tipo = {tipo: df.copy() for tipo, df in inv_z_ord.groupby("TipoRegalo", sort=False)}
        
        # 4. Iterar por cada tienda de la zona
        for idx_tienda, rowt in tiendas_z.iterrows():
            idt, tipo_tda = rowt["IDTienda"], rowt["TipoRegalo"]
            exito_tienda = False

            if tipo_tda in inv_por_tipo:
                ok, cods, descs, df_tipo_nuevo = intentar_asignar_para_tienda(inv_por_tipo[tipo_tda], n_regalos)
                if ok:
                    inv_por_tipo[tipo_tda] = df_tipo_nuevo
                    tdas.loc[idx_tienda, "REGALO_1"] = cods[0]
                    if len(cods) > 1:
                        tdas.loc[idx_tienda, "REGALO_2"] = cods[1]
                    tdas.loc[idx_tienda, "NOTAS"] = ""
                    exito_tienda = True
            
            if not exito_tienda and n_regalos > 1:
                if tipo_tda in inv_por_tipo:
                    ok, cods, descs, df_tipo_nuevo = intentar_asignar_para_tienda(inv_por_tipo[tipo_tda], 1)
                    if ok:
                        inv_por_tipo[tipo_tda] = df_tipo_nuevo
                        tdas.loc[idx_tienda, "REGALO_1"] = cods[0]
                        tdas.loc[idx_tienda, "NOTAS"] = f"Asignación parcial (1 de {n_regalos} solicitados)"
                        exito_tienda = True

            if not exito_tienda:
                excepciones.append({"IDTienda": idt, "NombreTienda": rowt["NombreTienda"], "Zona": rowt["Zona"], "Motivo": f"No hay stock suficiente para asignar los {n_regalos} regalos solicitados."})

        # 5. Consolidar cambios de inventario de la zona
        for tipo, df_tipo in inv_por_tipo.items():
            for idx, row in df_tipo.iterrows():
                inv_actualizado.loc[idx, "CantidadDisponible"] = row["CantidadDisponible"]

    # 6. Preparar salida
    df_tiendas_final = tdas
    df_inv_restante = inv_actualizado[inv_actualizado["CantidadDisponible"] > 0].copy()

    total_regalos = df_tiendas_final["REGALO_1"].ne("").sum() + df_tiendas_final["REGALO_2"].ne("").sum()
    tiendas_asignadas = df_tiendas_final["REGALO_1"].ne("").sum()
    
    reporte = [
        "==== REPORTE DE EJECUCIÓN ====",
        f"Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"EstrategiaDePriorizacion: {estrategia}",
        f"NumeroRegalosPorTienda: {n_regalos}",
        f"Tiendas procesadas: {len(tdas)}",
        f"Tiendas con asignación: {tiendas_asignadas}",
        f"Total de regalos asignados: {int(total_regalos)}",
        "\n---- Excepciones ----"
    ]
    if excepciones:
        for e in excepciones:
            reporte.append(f"[{e['Zona']}] {e['IDTienda']} - {e['NombreTienda']}: {e['Motivo']}")
    else:
        reporte.append("Sin excepciones.")

    # 7. Exportar a Excel en memoria con formato
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_tiendas_final.to_excel(writer, index=False, sheet_name="Asignacion")
        hoja_asig = writer.sheets['Asignacion']

        for column in hoja_asig.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except (TypeError, IndexError):
                    pass
            adjusted_width = (max_length + 2)
            hoja_asig.column_dimensions[column_letter].width = adjusted_width

        header_fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in hoja_asig[1]:
            cell.fill = header_fill
            cell.font = header_font

        df_inv_restante.to_excel(writer, index=False, sheet_name="InventarioRestante")

    excel_bytes = output.getvalue()

    return df_tiendas_final, df_inv_restante, "\n".join(reporte), excel_bytes