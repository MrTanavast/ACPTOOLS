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
