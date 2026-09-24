# Pruebas automáticas (LibreOffice)

El código VBA del núcleo se ejecuta de verdad en LibreOffice Calc (modo compatible
con VBA) para comprobar que genera exactamente lo mismo que el fichero de referencia.

```bash
python3 tests/run_all.py
```

- `test_facturas_lo.py`: CSV real de Contapluz (01692) → debe salir **byte a byte**
  igual que `datos/SUENLACE_referencia_01692.DAT` (el que se importó bien en a3).
- `test_diario_lo.py`: diario con casos límite → se compara con `oraculo_diario.py`
  (implementación independiente del registro tipo 0 en Python).
- `test_general_lo.py`: casos límite del perfil **GENERAL** de facturas emitidas (17 escenarios
  en una sola sesión de LibreOffice) comparados documento a documento con `oraculo_facturas.py`
  (implementación independiente de los registros 1/2 y 9 a partir de las posiciones de la
  especificación de a3 y de las reglas del procedimiento). Datos en `datos/general_*.csv`
  (`general_casos.csv` está en ANSI 1252 + CRLF, como lo guarda Excel en español) y un origen
  "hoja de Excel" con valores con tipo (fechas, números, % con formato %).
  Cubre: dos tipos de IVA y dos cuentas de ventas por línea, ticket, abono en negativo y en
  positivo, recargo (con y sin %), retención, cuentas con punto, cuenta de longitud incorrecta,
  IVA sin configurar, descuadre, línea a cero, anulada, subtipo 02, importes `1.234,56` y
  `1234.56`, fechas `15/01/2026`, `2026-01-15` y `15-01-26`, nº de más de 10 caracteres y los
  filtros de fecha, rango de documentos, series (incluir / excluir) y tipos.
  Mientras haya errores del VBA abiertos la prueba falla y marca cada diferencia con su
  hallazgo (`casos-limite-01` … `06`, lista `CONOCIDOS`); una diferencia sin documentar sale
  como `NUEVO`.
- `lo_merge.py` une los módulos en uno (LibreOffice no comparte `Type` entre módulos).
- `lo_bisect.py`, `lo_traza.py`: localizan errores de compilación / ejecución.

## Reglas para que el código funcione igual en Excel y en LibreOffice

1. Nada de `If ... Then ReDim ...` en una sola línea (LibreOffice borra la matriz): usar bloque.
2. Nada de `If ... Then Sub args Else ...` en una línea: usar bloque.
3. No asignar tipos (`Type`) completos (`a = b`): en LibreOffice se comparten por referencia.
   Rellenar campo a campo; para ordenar, ordenar índices.
4. No usar `Base` como nombre de campo o variable (palabra reservada en LibreOffice).
5. No pasar miembros de `With` (`.Campo`) como argumento `ByRef`.
6. No usar sufijos de tipo en literales (`@`), ni formatos regionales (`Format`, `CDbl`,
   `CStr` sobre decimales): los importes se construyen a mano.
7. Los `Type` no pueden contener matrices: usar matrices globales del módulo.
8. Para leer una `Collection` usar `LeerColeccion` (en LibreOffice una clave que no existe
   deja un objeto en la variable en lugar de dejarla vacía).
9. Un parámetro no puede llamarse igual que su procedimiento.
