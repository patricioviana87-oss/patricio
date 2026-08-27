# Conciliación de cheques / pagos pendientes

Programa que reemplaza el chequeo manual entre:

1. **Query de la cuenta** (mayor contable de "CHEQ EMITIDOS PEND DEB BCO", exportado
   como `.xls` desde el sistema — en realidad es texto separado por TABs).
2. **Informe de Pagos** (listado de egresos de valores con estado de conciliación,
   mismo formato de exportación).

Para cada egreso emitido en la Query, busca su fila correspondiente en el Informe de
Pagos y determina el estado:

- **CONCILIADO**: el Informe de Pagos indica `CONCILIADO = SI`.
- **PENDIENTE**: no está conciliado todavía → compone el saldo pendiente de la cuenta.
- **ANULADO**: el comprobante fue anulado (aparece un "Egreso de Valores Anulado" en
  la Query que lo cancela).
- **SIN INFORME DE PAGOS**: el egreso está en la Query pero no aparece en el Informe
  de Pagos cargado → hay que revisarlo a mano (por ejemplo, si el Informe no cubre
  la fecha de ese egreso).

También arma automáticamente el detalle de **Proveedor / Detalle / Monto / Fecha**
que antes se armaba a mano con el Informe de Pagos.

El cruce entre los dos archivos se hace por el **número de egreso** (columna
`NUMEROTRANSACCIONORIGEN` de la Query vs. `NRO_EGRESO_VALOR` del Informe de Pagos),
no por texto. Por eso no importa que el campo `NUMCONCIL` del Informe de Pagos venga
a veces como `Conc.Valor Nro.` (abreviado) y no como `Conciliación de Valores Nro.`
(como aparece en la Query): esa variación de formato no afecta el resultado, el
programa nunca compara esos textos entre sí.

## Instalación (una sola vez)

Necesitás Python 3 instalado. Después, desde la carpeta `conciliacion/`:

```bash
pip install -r requirements.txt
```

## Cómo correrlo cada mes

```bash
python3 conciliar.py QUERY_CUENTA.xls INFORME_PAGOS.xls SALIDA.xlsx
```

- **QUERY_CUENTA.xls**: el archivo que bajás del mayor de la cuenta (el "query de
  cheques") con el movimiento del mes que estás conciliando.
- **INFORME_PAGOS.xls**: el Informe de Pagos que bajás de intranet (conviene que
  cubra varios meses hacia atrás, para que aparezcan conciliaciones de egresos
  emitidos antes).
- **SALIDA.xlsx**: el nombre que le querés poner al Excel de resultado (se crea o
  se pisa si ya existe).

Ejemplo real, mes a mes:

```bash
python3 conciliar.py Query_Cheques_210109_JUL.xls Informe_de_Pagos_JUL.xls conciliacion_julio.xlsx
python3 conciliar.py Query_Cheques_210109_AGO.xls Informe_de_Pagos_AGO.xls conciliacion_agosto.xlsx
python3 conciliar.py Query_Cheques_210109_SEP.xls Informe_de_Pagos_SEP.xls conciliacion_septiembre.xlsx
```

No hace falta pisar el archivo del mes anterior por el nuevo, ni pegar nada a mano:
cada corrida agrega lo nuevo a lo ya cargado (ver "Acumulación" abajo). Los archivos
de Query e Informe de cada mes los podés dejar donde quieras (Descargas, una carpeta
del mes, etc.) — solo tenés que apuntar la ruta correcta al llamarlos.

### Acumulación mes a mes (para que los pendientes que se saldan después se reflejen)

Cada vez que corrés el programa, **acumula** lo que le cargaste (Query + Informe de
Pagos) en una carpeta `historial/` que se crea automáticamente al lado de
`conciliar.py`, y concilia contra **todo lo acumulado hasta ese momento**, no solo
contra el archivo de ese día. Así, un pago que en julio queda pendiente y se concilia
recién en agosto, en la corrida de agosto va a aparecer correctamente como
CONCILIADO — sin que tengas que pegar nada a mano.

- No se duplica si volvés a cargar el mismo mes dos veces.
- Si el Informe de Pagos de un mes actualiza el estado de un egreso viejo (pasó de
  "NO" a "SI" conciliado), el historial se actualiza con el estado más nuevo.
- La carpeta `historial/` **no se sube al repositorio** (está en `.gitignore`,
  contiene datos contables y de proveedores). No la borres entre corridas, es la
  que le da memoria al programa. Si querés hacerle una copia de resguardo, es solo
  copiar la carpeta.
- Si preferís guardar el historial en otro lado (por ejemplo, una carpeta
  compartida), pasala como cuarto parámetro:
  ```bash
  python3 conciliar.py Query.xls Informe.xls salida.xlsx /ruta/a/mi/historial
  ```

## El Excel de salida

Tiene 5 hojas:

- **Resumen**: cantidad de egresos por estado, saldo pendiente total y control
  contra el movimiento acumulado en el historial.
- **Detalle**: una fila por egreso (de todo el historial acumulado), con Proveedor /
  Detalle / Monto / Fecha / Estado y los datos de conciliación (fecha, número de
  lote, CUIT, banco). Coloreado por estado (verde = conciliado, amarillo =
  pendiente, gris = anulado, rojo = a revisar). Tiene autofiltro, así que podés
  filtrar por Estado = PENDIENTE para ver solo lo que falta conciliar.
- **Por Proveedor**: total pendiente agrupado por proveedor.
- **Revisar**: egresos sin cruce contra el Informe de Pagos, o con un importe
  distinto entre la Query y el Informe (posible error de carga).
- **Fuera de Periodo**: pagos del Informe que no corresponden a ningún egreso
  cargado hasta ahora (normal si el Informe cubre más meses de los que ya
  cargaste — no requiere acción).

## Notas importantes

- El **saldo pendiente** se va a acercar cada vez más al saldo contable real de la
  cuenta a medida que pasan los meses y acumulás más historial. Las primeras
  corridas pueden no cerrar exacto (hay conciliaciones que cancelan egresos
  emitidos antes de la primera Query que cargaste) — es esperable, y se explica
  también en la hoja Resumen del Excel.
- Si el Informe de Pagos trae dos filas para el mismo egreso (una vieja "NO
  conciliado" y una nueva "SI"), el programa usa la fila conciliada.
- Los números vienen en formato argentino (coma decimal) y los archivos en
  codificación Latin-1; el programa ya lo maneja.
