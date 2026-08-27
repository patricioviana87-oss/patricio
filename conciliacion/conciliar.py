#!/usr/bin/env python3
"""
Conciliación automática: Query de la cuenta (mayor contable) vs. Informe de Pagos.

Uso:
    python3 conciliar.py QUERY_CUENTA.xls INFORME_PAGOS.xls SALIDA.xlsx [carpeta_historial]

Cada corrida ACUMULA la Query y el Informe de Pagos que le vayas cargando (mes a
mes) en `carpeta_historial` (por defecto: `historial/`, junto al script), y
concilia contra TODO lo acumulado hasta ahora — no solo contra el archivo del
día. Así, un pago que queda pendiente este mes y se concilia recién el mes que
viene queda bien reflejado la próxima vez que corras el programa. Ver README.md.

Lógica (replica el criterio manual, ver instrucciones del archivo de control):
  - Cada "Egreso de Valores Nro. X" en el mayor es un pago emitido (pasa a
    pendiente). Si en algún momento el banco lo debita, aparece un DEBE que lo
    cancela: o bien dentro de un lote "Conciliación de Valores Nro. Y", o bien
    como "Egreso de Valores Anulado Nro. X" (anulación del comprobante).
  - El Informe de Pagos ya indica, por cada Egreso, si está CONCILIADO (SI/NO),
    con qué lote (NUMCONCIL) y en qué fecha. Se usa como fuente de verdad del
    estado, y se valida contra el mayor. El campo NUMCONCIL puede venir
    abreviado ("Conc.Valor Nro.") en vez del texto completo que usa el mayor
    ("Conciliación de Valores Nro."); no importa, porque el cruce NO se hace
    por ese texto sino por el número de egreso (NUMEROTRANSACCIONORIGEN vs.
    NRO_EGRESO_VALOR), así que esa variación de formato no afecta el resultado.
  - Saldo pendiente = suma de los egresos NO conciliados y no anulados.
    Ese saldo pendiente debe coincidir con el total acumulado del mayor de la
    cuenta (columna TOTAL/HABER-DEBE) — si no coincide, hay diferencias para
    revisar a mano, y el programa las lista aparte.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from historial import actualizar_historial_informe, actualizar_historial_query
from parsers import leer_query_cuenta, leer_informe_pagos
from reporte import escribir_reporte

TOLERANCIA = 0.5  # pesos; margen de redondeo para considerar "conciliado en el mayor"

HISTORIAL_POR_DEFECTO = Path(__file__).parent / "historial"


def conciliar(q: pd.DataFrame, informe: pd.DataFrame) -> dict:
    """Concilia a partir de las tablas YA ACUMULADAS de Query e Informe de Pagos."""

    # El Informe de Pagos a veces trae más de una fila para el mismo egreso
    # (p.ej. una versión vieja "NO conciliado" y una más nueva "SI"). Nos
    # quedamos con la fila conciliada si existe alguna; si no, con la primera.
    informe = informe.sort_values("CONCILIADO", ascending=False).drop_duplicates(
        subset="NRO_EGRESO_VALOR", keep="first"
    )

    egresos = q[q["TIPO_MOVIMIENTO"] == "EGRESO"].copy()
    anulados = q[q["TIPO_MOVIMIENTO"] == "EGRESO_ANULADO"].copy()
    nros_anulados = set(anulados["NUMEROTRANSACCIONORIGEN"].dropna().tolist())

    # El universo a conciliar lo define la Query (los egresos emitidos en el
    # período que bajaste del mayor). El Informe de Pagos puede cubrir un
    # rango de fechas más amplio (p.ej. "últimos 3 meses") a propósito, para
    # poder encontrar la conciliación de egresos emitidos antes: por eso el
    # cruce es "left" desde la Query, no un outer join.
    detalle = egresos.merge(
        informe,
        left_on="NUMEROTRANSACCIONORIGEN",
        right_on="NRO_EGRESO_VALOR",
        how="left",
        suffixes=("_MAYOR", "_INFORME"),
        indicator=True,
    )

    # Informe de Pagos que no corresponden a ningún egreso de esta Query:
    # no son un error, son pagos fuera del período que estás conciliando.
    nros_query = set(egresos["NUMEROTRANSACCIONORIGEN"].dropna().tolist())
    fuera_de_periodo = informe[~informe["NRO_EGRESO_VALOR"].isin(nros_query)].copy()

    def estado_fila(row) -> str:
        nro = row["NUMEROTRANSACCIONORIGEN"]
        if pd.notna(nro) and nro in nros_anulados:
            return "ANULADO"
        if row["_merge"] == "left_only":
            return "SIN INFORME DE PAGOS"
        if row.get("CONCILIADO") == "SI":
            return "CONCILIADO"
        return "PENDIENTE"

    detalle["ESTADO"] = detalle.apply(estado_fila, axis=1)

    detalle["NRO_EGRESO"] = detalle["NUMEROTRANSACCIONORIGEN"].combine_first(detalle["NRO_EGRESO_VALOR"])
    detalle["PROVEEDOR"] = detalle["DESTINATARIO"]
    detalle["DETALLE"] = detalle["DETALLE_OP"].combine_first(detalle["TRANSACCIONORIGEN"])
    detalle["MONTO"] = detalle["IMPORTE"].combine_first(detalle["MONTO_SIGNADO"].abs())
    detalle["FECHA"] = detalle["FECHAVENC"].combine_first(detalle["FECHATR"])

    diff_monto = (detalle["IMPORTE"] - detalle["MONTO_SIGNADO"].abs()).abs()
    detalle["DIFERENCIA_IMPORTE"] = diff_monto.where(diff_monto > TOLERANCIA)

    columnas_salida = [
        "NRO_EGRESO",
        "PROVEEDOR",
        "DETALLE",
        "MONTO",
        "FECHA",
        "ESTADO",
        "CONCILIADO",
        "FECHA_CONCILIACION",
        "NUMCONCIL",
        "CUIT",
        "BANCO",
        "DIFERENCIA_IMPORTE",
    ]
    detalle_final = detalle[columnas_salida].sort_values(["ESTADO", "FECHA"], na_position="last")

    saldo_pendiente = float(detalle.loc[detalle["ESTADO"] == "PENDIENTE", "MONTO"].sum())
    total_mayor = float(q["MONTO_SIGNADO"].sum())
    diferencia_control = saldo_pendiente - total_mayor

    resumen = {
        "total_egresos": int((q["TIPO_MOVIMIENTO"] == "EGRESO").sum()),
        "conciliados": int((detalle["ESTADO"] == "CONCILIADO").sum()),
        "pendientes": int((detalle["ESTADO"] == "PENDIENTE").sum()),
        "anulados": int((detalle["ESTADO"] == "ANULADO").sum()),
        "sin_informe": int((detalle["ESTADO"] == "SIN INFORME DE PAGOS").sum()),
        "pagos_fuera_de_periodo": int(len(fuera_de_periodo)),
        "saldo_pendiente": saldo_pendiente,
        "total_mayor_cuenta": total_mayor,
        "diferencia_control": diferencia_control,
    }

    por_proveedor = (
        detalle.loc[detalle["ESTADO"] == "PENDIENTE"]
        .groupby("PROVEEDOR", dropna=False)["MONTO"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    revisar = detalle_final[
        detalle_final["ESTADO"].isin(["SIN INFORME DE PAGOS"]) | detalle_final["DIFERENCIA_IMPORTE"].notna()
    ]

    fuera_de_periodo_out = fuera_de_periodo[
        ["NRO_EGRESO_VALOR", "DESTINATARIO", "DETALLE_OP", "IMPORTE", "FECHAVENC", "CONCILIADO"]
    ].rename(
        columns={
            "NRO_EGRESO_VALOR": "NRO_EGRESO",
            "DESTINATARIO": "PROVEEDOR",
            "DETALLE_OP": "DETALLE",
            "IMPORTE": "MONTO",
            "FECHAVENC": "FECHA",
        }
    )

    return {
        "detalle": detalle_final,
        "resumen": resumen,
        "por_proveedor": por_proveedor,
        "revisar": revisar,
        "fuera_de_periodo": fuera_de_periodo_out,
    }


def main():
    if len(sys.argv) not in (4, 5):
        print(__doc__)
        sys.exit(1)
    path_query, path_informe, path_salida = sys.argv[1:4]
    historial_dir = Path(sys.argv[4]) if len(sys.argv) == 5 else HISTORIAL_POR_DEFECTO

    qc_nuevo = leer_query_cuenta(path_query)
    informe_nuevo = leer_informe_pagos(path_informe)

    q_acumulada = actualizar_historial_query(qc_nuevo.df, historial_dir)
    informe_acumulado = actualizar_historial_informe(informe_nuevo, historial_dir)

    resultado = conciliar(q_acumulada, informe_acumulado)
    escribir_reporte(resultado, path_salida)

    print(f"Historial acumulado en:    {historial_dir}")
    r = resultado["resumen"]
    print(f"Egresos analizados (histórico total): {r['total_egresos']}")
    print(f"  Conciliados:             {r['conciliados']}")
    print(f"  Pendientes:              {r['pendientes']}")
    print(f"  Anulados:                {r['anulados']}")
    print(f"  Sin informe de pagos:    {r['sin_informe']}")
    print(f"  Pagos fuera de período:  {r['pagos_fuera_de_periodo']} (informativo, no es un error)")
    print()
    print(f"Saldo pendiente (Informe): {r['saldo_pendiente']:,.2f}")
    print(f"Movimiento neto acumulado en historial: {r['total_mayor_cuenta']:,.2f}")
    print(f"Diferencia:                {r['diferencia_control']:,.2f}")
    print()
    print(f"Reporte generado en: {path_salida}")


if __name__ == "__main__":
    main()
