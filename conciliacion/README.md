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

## Uso

```bash
pip install -r requirements.txt
python3 conciliar.py QUERY_CUENTA.xls INFORME_PAGOS.xls SALIDA.xlsx
```

Ejemplo:

```bash
python3 conciliar.py Query_Cheques_210109.xls Informe_de_Pagos_1049400.xls conciliacion_julio.xlsx
```

El Excel de salida tiene 5 hojas:

- **Resumen**: cantidad de egresos por estado, saldo pendiente total y control contra
  el movimiento de la Query cargada.
- **Detalle**: una fila por egreso, con Proveedor / Detalle / Monto / Fecha / Estado
  y los datos de conciliación (fecha, número de lote, CUIT, banco). Coloreado por
  estado (verde = conciliado, amarillo = pendiente, gris = anulado, rojo = a revisar).
- **Por Proveedor**: total pendiente agrupado por proveedor.
- **Revisar**: egresos sin cruce contra el Informe de Pagos, o con un importe distinto
  entre la Query y el Informe (posible error de carga).
- **Fuera de Periodo**: pagos del Informe que no corresponden a ningún egreso de la
  Query cargada (normal si el Informe cubre más meses que la Query — no requiere acción).

## Notas importantes

- El **saldo pendiente** solo va a coincidir exactamente con el saldo contable real de
  la cuenta si la Query cargada tiene el **historial completo** de la cuenta (todos los
  meses acumulados), igual que en el proceso manual. Si cargás un solo mes, la
  diferencia contra el "movimiento neto de la Query" es esperable (hay conciliaciones
  de ese mes que cancelan egresos de meses anteriores).
- Si el Informe de Pagos trae dos filas para el mismo egreso (una vieja "NO
  conciliado" y una nueva "SI"), el programa usa la fila conciliada.
- Los números vienen en formato argentino (coma decimal) y los archivos en codificación
  Latin-1; el programa ya lo maneja.
