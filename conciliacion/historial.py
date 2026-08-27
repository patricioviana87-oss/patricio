"""
Acumula la Query de la cuenta y el Informe de Pagos entre corridas del programa,
para poder conciliar contra el HISTÓRICO completo (no solo el mes que acabás de
cargar) — así los pendientes de un mes que se saldan al mes siguiente quedan bien
reflejados, igual que en el proceso manual (que va pegando mes a mes en la misma
planilla).

Se guardan dos archivos CSV en la carpeta de historial (por defecto
`conciliacion/historial/`, junto al script):

  - query_historial.csv    -> movimientos del mayor de la cuenta (todos los meses)
  - informe_historial.csv  -> egresos del informe de pagos (todos los meses)

Estos CSV son datos contables/de proveedores: NO se suben al repositorio
(están en .gitignore). Hacer una copia de resguardo si te importa no perderlos.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

QUERY_COLUMNS = [
    "NUMERO",
    "NUMEROITEM",
    "TRANSACCIONORIGEN",
    "NUMEROTRANSACCIONORIGEN",
    "TIPO_MOVIMIENTO",
    "FECHATR",
    "DEBE",
    "HABER",
    "MONTO_SIGNADO",
]

INFORME_COLUMNS = [
    "NRO_EGRESO_VALOR",
    "DESTINATARIO",
    "IMPORTE",
    "FECHAVENC",
    "BANCO",
    "DETALLE_OP",
    "CONCILIADO",
    "FECHA_CONCILIACION",
    "NUMCONCIL",
    "CUIT",
    "CBU",
]

QUERY_FILE = "query_historial.csv"
INFORME_FILE = "informe_historial.csv"


def _leer_csv_si_existe(path: Path, columnas_fecha: list[str], columnas_enteras: list[str]) -> pd.DataFrame | None:
    if not path.exists():
        return None
    df = pd.read_csv(path)
    for col in columnas_fecha:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    for col in columnas_enteras:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
    return df


def actualizar_historial_query(nuevo: pd.DataFrame, historial_dir: Path) -> pd.DataFrame:
    historial_dir.mkdir(parents=True, exist_ok=True)
    path = historial_dir / QUERY_FILE
    nuevo = nuevo[QUERY_COLUMNS].copy()

    previo = _leer_csv_si_existe(path, ["FECHATR"], ["NUMEROTRANSACCIONORIGEN", "NUMERO", "NUMEROITEM"])
    combinado = pd.concat([nuevo, previo], ignore_index=True) if previo is not None else nuevo

    # Un mismo asiento/ítem (NUMERO + NUMEROITEM) es único: si volvés a cargar
    # la Query de un mes que ya estaba, no se duplica.
    combinado = combinado.drop_duplicates(subset=["NUMERO", "NUMEROITEM"], keep="first")
    combinado["NUMEROTRANSACCIONORIGEN"] = pd.to_numeric(
        combinado["NUMEROTRANSACCIONORIGEN"], errors="coerce"
    ).astype("Int64")
    combinado.to_csv(path, index=False)
    return combinado


def actualizar_historial_informe(nuevo: pd.DataFrame, historial_dir: Path) -> pd.DataFrame:
    historial_dir.mkdir(parents=True, exist_ok=True)
    path = historial_dir / INFORME_FILE
    nuevo = nuevo[INFORME_COLUMNS].copy()

    previo = _leer_csv_si_existe(path, ["FECHAVENC", "FECHA_CONCILIACION"], ["NRO_EGRESO_VALOR"])
    combinado = pd.concat([nuevo, previo], ignore_index=True) if previo is not None else nuevo

    # Para un mismo egreso puede haber varias filas entre corridas distintas
    # (el Informe de Pagos suele traer los últimos 3 meses, así que un mismo
    # egreso reaparece; y a veces trae una fila vieja "NO conciliado" junto a
    # una nueva "SI"). Nos quedamos con la conciliada si existe alguna.
    combinado = combinado.sort_values("CONCILIADO", ascending=False).drop_duplicates(
        subset="NRO_EGRESO_VALOR", keep="first"
    )
    combinado["NRO_EGRESO_VALOR"] = pd.to_numeric(combinado["NRO_EGRESO_VALOR"], errors="coerce").astype("Int64")
    combinado.to_csv(path, index=False)
    return combinado
