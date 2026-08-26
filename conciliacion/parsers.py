"""
Lectura de los dos archivos fuente para la conciliación bancaria de cheques/pagos:

1) Query de la cuenta (mayor contable de la cuenta de "cheques emitidos
   pendientes de débito bancario"), exportado como .xls pero en realidad es
   texto separado por TABs, con unas líneas de metadata al principio.

2) Informe de Pagos (listado de egresos de valores con su estado de
   conciliación), mismo formato: texto separado por TABs con metadata previa.

Ambos archivos vienen con codificación latin-1 y números con coma decimal
(formato argentino), así que se normalizan acá.
"""
from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass

import pandas as pd

ENCODING = "latin-1"


def _find_header_row(lines: list[str], marker: str) -> int:
    """Devuelve el índice (0-based) de la línea que contiene los nombres de columna."""
    for i, line in enumerate(lines):
        if marker in line:
            return i
    raise ValueError(f"No se encontró la fila de encabezados (buscando {marker!r})")


def _read_tsv_ragged(path: str, header_idx: int) -> pd.DataFrame:
    """
    Lee un archivo TSV a partir de la fila de encabezado `header_idx`, tolerando
    filas con columnas de más o de menos (por tabs "sueltos" dentro de texto
    libre, como el campo DETALLE_OP del Informe de Pagos).
    """
    with open(path, encoding=ENCODING, newline="") as f:
        lines = f.readlines()

    header_line = lines[header_idx].rstrip("\r\n")
    columns = header_line.split("\t")
    # el export suele dejar una columna vacía al final (tab colgante)
    n_cols = len(columns)
    if columns and columns[-1] == "":
        n_cols -= 1
        columns = columns[:-1]

    reader = csv.reader(lines[header_idx + 1 :], delimiter="\t")
    rows = []
    for raw in reader:
        if not raw or all(cell == "" for cell in raw):
            continue
        if len(raw) == n_cols or len(raw) == n_cols + 1:
            rows.append(raw[:n_cols])
        elif len(raw) > n_cols:
            # Un tab de más "de sobra" dentro de un campo de texto libre:
            # se pega el token extra al campo anterior (best effort).
            extra = len(raw) - n_cols
            fixed = raw[: n_cols - 1] + [" ".join(raw[n_cols - 1 : n_cols - 1 + extra + 1])] + raw[n_cols - 1 + extra + 1 :]
            rows.append(fixed[:n_cols])
        else:
            rows.append(raw + [""] * (n_cols - len(raw)))

    return pd.DataFrame(rows, columns=columns)


def _to_float_ar(series: pd.Series) -> pd.Series:
    """Convierte strings con coma decimal (formato AR) a float. Vacío -> NaN."""
    cleaned = (
        series.astype(str)
        .str.strip()
        .replace({"": None, "nan": None, "None": None})
    )
    cleaned = cleaned.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(cleaned, errors="coerce")


def _extraer_numero(texto: pd.Series) -> pd.Series:
    """Extrae el número de comprobante (dígitos) de un texto tipo 'Nro. 00075622'."""
    return texto.astype(str).str.extract(r"(\d+)$")[0].astype("Int64")


@dataclass
class QueryCuenta:
    df: pd.DataFrame
    total_cuenta: float


def leer_query_cuenta(path: str) -> QueryCuenta:
    with open(path, encoding=ENCODING) as f:
        lines = f.readlines()
    header_idx = _find_header_row(lines, "NUMEROTRANSACCIONORIGEN")
    df = _read_tsv_ragged(path, header_idx)

    df["DEBE"] = _to_float_ar(df["DEBE"])
    df["HABER"] = _to_float_ar(df["HABER"])
    df["TOTAL"] = _to_float_ar(df["TOTAL"])
    # Signo: HABER aumenta el pendiente (egreso emitido), DEBE lo cancela (conciliado/anulado)
    df["MONTO_SIGNADO"] = df["HABER"].fillna(0) - df["DEBE"].fillna(0)

    df["FECHATR"] = pd.to_datetime(df["FECHATR"], format="%d/%m/%Y", errors="coerce")
    df["FECHAAPLICACION"] = pd.to_datetime(df["FECHAAPLICACION"], format="%d/%m/%Y", errors="coerce")

    df["NUMEROTRANSACCIONORIGEN"] = pd.to_numeric(df["NUMEROTRANSACCIONORIGEN"], errors="coerce").astype("Int64")

    texto = df["TRANSACCIONORIGEN"].astype(str)
    df["TIPO_MOVIMIENTO"] = "OTRO"
    df.loc[texto.str.startswith("Egreso de Valores Anulado"), "TIPO_MOVIMIENTO"] = "EGRESO_ANULADO"
    df.loc[
        texto.str.startswith("Egreso de Valores Nro") & ~texto.str.contains("Anulado"),
        "TIPO_MOVIMIENTO",
    ] = "EGRESO"
    df.loc[texto.str.startswith("Conciliación de Valores"), "TIPO_MOVIMIENTO"] = "CONCILIACION_BATCH"

    total_cuenta = float(df["MONTO_SIGNADO"].sum())
    return QueryCuenta(df=df, total_cuenta=total_cuenta)


def leer_informe_pagos(path: str) -> pd.DataFrame:
    with open(path, encoding=ENCODING) as f:
        lines = f.readlines()
    header_idx = _find_header_row(lines, "NRO_EGRESO_VALOR")
    df = _read_tsv_ragged(path, header_idx)

    df["IMPORTE"] = _to_float_ar(df["IMPORTE"])
    df["NRO_EGRESO_VALOR"] = pd.to_numeric(df["NRO_EGRESO_VALOR"], errors="coerce").astype("Int64")
    df["FECHAVENC"] = pd.to_datetime(df["FECHAVENC"], format="%d/%m/%Y", errors="coerce")
    df["FECHA_CONCILIACION"] = pd.to_datetime(df["FECHA_CONCILIACION"], format="%d/%m/%Y", errors="coerce")
    df["CONCILIADO"] = df["CONCILIADO"].astype(str).str.strip().str.upper()
    df["NUMCONCIL_NUM"] = _extraer_numero(df["NUMCONCIL"])
    return df
