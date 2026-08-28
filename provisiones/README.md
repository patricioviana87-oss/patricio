# Asientos de provisiones

Programa que genera los asientos contables de provisiones (un Excel por
sector, listo para subir al sistema contable) a partir del archivo de
detalle de provisiones que arma cada sector.

## Archivos de entrada

1. **Detalle de provisiones** (`PROVISIONES.xlsx`): un Excel con la hoja
   `PROVISIONES`, con un renglón por gasto a provisionar. Columnas usadas:

   - `sector`
   - `Periodo de aplicación` (nombre del mes, ej. "Agosto")
   - `Proveedor/ Razon social`
   - `Descripción del gasto`
   - `Importe estimado en pesos` (puede venir como número o como texto con
     formato argentino: `$ 1.050.537,47` o `$ 782.808`)
   - `Cuenta` (cuenta contable de gasto)
   - `Centro de costo`
   - `Código de actividad`
   - `OC - RG - CF - RC`
   - `COMENTARIOS` (opcional, se vuelca a la columna `Nota` del asiento)

   Los renglones sin `sector` o con importe 0 se descartan. Si aparecen
   varios meses en el mismo archivo, se genera un asiento separado por cada
   combinación sector + mes.

## Qué genera

Por cada sector (y mes) arma un asiento con las mismas 14 columnas que usa
el archivo de asientos de origen (`Numero`, `Fecha`, `Detalle`, `Fecha
Aplicacion`, `C.Costos`, `Moneda`, `Cotizacion`, `Cuenta`, `Detalle Item`,
`Tipo (D|H)`, `Importe`, `Cod.Ejercicio`, `Codigo Actividad`, `Nota`):

- **Una línea por gasto** (columna `Tipo` = `D`): cuenta, centro de costo,
  código de actividad e importe tal como vienen en el detalle, con el
  detalle armado como `Provisión {mes} {año} - {Proveedor} - {Descripción}
  - {OC/RG/CF/RC}`.
- **Una línea de contrapartida por cada centro de costo** distinto que haya
  en el sector, contra la cuenta de provisión a pagar `220111`, con el
  importe en negativo (la suma de los gastos de ese centro de costo). Sigue
  el mismo criterio del archivo de origen: **todas** las líneas usan `Tipo
  D`; es el signo del importe el que indica debe/haber, no la columna Tipo.
- `Numero` = 200 (fijo) y `Fecha` = `Fecha Aplicacion` = último día del mes
  provisionado, porque los asientos siempre se suben con fecha de fin de
  mes.

Cada asiento queda balanceado (la suma de sus importes da 0): al terminar,
el programa imprime el total por sector para verificarlo rápido contra el
archivo de detalle.

## Instalación (una sola vez)

Necesitás Python 3. Desde la carpeta `provisiones/`:

```bash
pip install -r requirements.txt
```

## Cómo correrlo cada mes

```bash
python3 generar.py PROVISIONES.xlsx 2026
```

- **PROVISIONES.xlsx**: el archivo de detalle de provisiones del mes.
- **2026**: el año del ejercicio (se usa para `Fecha Aplicacion` y para
  `Cod.Ejercicio`, ej. `EJ2026`).
- Opcionalmente, un tercer argumento con la carpeta de salida (por defecto
  `provisiones/salida/`).

Se genera un archivo por sector, por ejemplo:

```
salida/08_Asiento_provision_FAJU.xlsx
salida/08_Asiento_provision_FACO.xlsx
salida/08_Asiento_provision_SAE.xlsx
...
```

## Supuestos a confirmar / ajustar si cambian

- La cuenta de contrapartida de provisión a pagar es siempre `220111`, para
  todos los sectores (`generar.py:CUENTA_PROVISION`).
- El código de actividad de la línea de contrapartida es siempre `9999`
  (`generar.py:CODIGO_ACTIVIDAD_CONTRA`).
- `Numero` de asiento fijo en `200` (`generar.py:NUMERO_ASIENTO`).

Si alguna de estas reglas cambia (por ejemplo, la cuenta de contrapartida
pasa a depender del sector), se ajusta en `generar.py`.
