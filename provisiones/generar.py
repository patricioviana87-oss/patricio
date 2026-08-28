#!/usr/bin/env python3
"""
Genera los asientos contables de provisiones, un archivo Excel por sector, a
partir del archivo de detalle de provisiones (hoja "PROVISIONES").

Uso:
    python3 generar.py PROVISIONES.xlsx ANIO [carpeta_salida]

Por cada sector (y mes, tomado de la columna "Periodo de aplicación") arma:

  - Una línea "D" (debe) por cada gasto a provisionar: la cuenta de gasto,
    el centro de costo, el código de actividad y el importe tal como vienen
    en el detalle de provisiones.
  - Una línea de contrapartida por cada centro de costo distinto que
    aparezca dentro del sector, contra la cuenta de provisión a pagar
    (220111), con el importe en negativo. Sigue el mismo criterio del
    archivo de asientos de origen: TODAS las líneas usan Tipo "D", y es el
    signo del importe el que indica si es debe o haber (no la columna Tipo).

Todas las líneas de un mismo asiento comparten Numero, Fecha y Fecha
Aplicación = último día del mes provisionado, porque los asientos siempre se
suben con fecha de fin de mes.
"""
from __future__ import annotations

import sys
from calendar import monthrange
from collections import defaultdict
from datetime import date
from pathlib import Path

from parsers import MESES_ES, leer_provisiones
from reporte import Linea, escribir_asiento

NUMERO_ASIENTO = 200
CUENTA_PROVISION = "220111"
CODIGO_ACTIVIDAD_CONTRA = 9999
MONEDA = "Pesos"

SALIDA_POR_DEFECTO = Path(__file__).parent / "salida"


def _sanitizar_nombre(sector: str) -> str:
    return sector.strip().replace(" ", "_").replace(".", "").replace("/", "-")


def generar_asientos(path_provisiones: str, anio: int, carpeta_salida: Path, on_linea=print) -> list[Path]:
    provisiones = leer_provisiones(path_provisiones)
    if not provisiones:
        raise ValueError("El archivo de provisiones no tiene renglones para procesar")

    por_sector_mes: dict[tuple[str, int], list] = defaultdict(list)
    for p in provisiones:
        por_sector_mes[(p.sector, p.mes)].append(p)

    carpeta_salida.mkdir(parents=True, exist_ok=True)
    archivos = []

    for (sector, mes), items in sorted(por_sector_mes.items()):
        ultimo_dia = date(anio, mes, monthrange(anio, mes)[1])
        mes_nombre = MESES_ES[mes]
        detalle_header = f"PROVISION DE GASTOS {mes:02d}/{anio} {sector}"
        cod_ejercicio = f"EJ{anio}"

        lineas: list[Linea] = []
        totales_por_cc: dict[str, float] = defaultdict(float)

        for p in items:
            detalle_item = f"Provisión {mes_nombre} {anio} - {p.proveedor} - {p.descripcion} - {p.referencia}"
            lineas.append(
                Linea(
                    numero=NUMERO_ASIENTO,
                    fecha=ultimo_dia,
                    detalle=detalle_header,
                    fecha_aplicacion=ultimo_dia,
                    centro_costo=p.centro_costo,
                    moneda=MONEDA,
                    cotizacion=1,
                    cuenta=p.cuenta,
                    detalle_item=detalle_item,
                    tipo="D",
                    importe=round(p.importe, 2),
                    cod_ejercicio=cod_ejercicio,
                    codigo_actividad=p.codigo_actividad,
                    nota=p.comentario,
                )
            )
            totales_por_cc[p.centro_costo] += p.importe

        for centro_costo, total in totales_por_cc.items():
            lineas.append(
                Linea(
                    numero=NUMERO_ASIENTO,
                    fecha=ultimo_dia,
                    detalle=detalle_header,
                    fecha_aplicacion=ultimo_dia,
                    centro_costo=centro_costo,
                    moneda=MONEDA,
                    cotizacion=1,
                    cuenta=CUENTA_PROVISION,
                    detalle_item=detalle_header,
                    tipo="D",
                    importe=round(-total, 2),
                    cod_ejercicio=cod_ejercicio,
                    codigo_actividad=CODIGO_ACTIVIDAD_CONTRA,
                    nota=None,
                )
            )

        nombre_archivo = f"{mes:02d}_Asiento_provision_{_sanitizar_nombre(sector)}.xlsx"
        path_salida = carpeta_salida / nombre_archivo
        escribir_asiento(lineas, path_salida)
        archivos.append(path_salida)

        total_sector = sum(p.importe for p in items)
        on_linea(f"  {sector} ({mes:02d}/{anio}): {len(items)} líneas, total ${total_sector:,.2f} -> {nombre_archivo}")

    return archivos


def main() -> None:
    if len(sys.argv) < 3:
        print("Uso: python3 generar.py PROVISIONES.xlsx ANIO [carpeta_salida]")
        sys.exit(1)

    path_provisiones = sys.argv[1]
    anio = int(sys.argv[2])
    carpeta_salida = Path(sys.argv[3]) if len(sys.argv) > 3 else SALIDA_POR_DEFECTO

    print(f"Generando asientos de provisiones ({path_provisiones})...")
    archivos = generar_asientos(path_provisiones, anio, carpeta_salida)
    print(f"\nListo: {len(archivos)} archivo(s) en {carpeta_salida}/")


if __name__ == "__main__":
    main()
