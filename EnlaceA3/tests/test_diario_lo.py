"""Prueba del motor de diarios en LibreOffice contra el oráculo Python."""
import pathlib, sys
from lo_merge import merge
import lo_runner, oraculo_diario
T = pathlib.Path(__file__).resolve().parent
CSV = T / "datos" / "diario_casos.csv"
OUT = T / "salida" / "SUENLACE_lo_diario.DAT"
CSV2 = T / "datos" / "diario_punto.csv"
OUT2 = T / "salida" / "SUENLACE_lo_diario2.DAT"
OUT.parent.mkdir(exist_ok=True)
TEST = r'''
Function TestDiario() As String
    On Error GoTo EH
    Dim emp As TEmpresa, fil As TFiltros, res As TResumen, msg As String, datos As Variant, s As String, i As Long
    emp.Codigo = "01696": emp.Digitos = 9
    fil.ExcluirDescuadrados = True
    datos = ParsearCSV(LeerFicheroTexto("__CSV__"), "auto")
    If Not DiarioProcesar(emp, fil, datos, res, msg) Then TestDiario = "ERR " & msg: Exit Function
    EscribirDat "__OUT__"
    s = "asientos=" & res.UnidadesLeidas & " exp=" & res.UnidadesExportadas & " excl=" & res.UnidadesExcluidas & _
        " lineas=" & gNDat & " debe=" & ImporteTexto(res.Debe) & " haber=" & ImporteTexto(res.Haber)
    For i = 1 To gNInc
        s = s & Chr(10) & gInc(i).Gravedad & " | " & gInc(i).Referencia & " | fila " & gInc(i).Fila & " | " & gInc(i).Texto
    Next i
    TestDiario = s
    Exit Function
EH:
    TestDiario = "ERR " & Err.Number & " " & Err.Description
End Function

Function TestDiario2() As String
    On Error GoTo EH
    Dim emp As TEmpresa, fil As TFiltros, res As TResumen, msg As String, datos As Variant, s As String, i As Long
    emp.Codigo = "01696": emp.Digitos = 9
    fil.ExcluirDescuadrados = False
    datos = ParsearCSV(LeerFicheroTexto("__CSV2__"), "auto")
    If Not DiarioProcesar(emp, fil, datos, res, msg) Then TestDiario2 = "ERR " & msg: Exit Function
    EscribirDat "__OUT2__"
    s = "asientos=" & res.UnidadesLeidas & " exp=" & res.UnidadesExportadas & " excl=" & res.UnidadesExcluidas
    For i = 1 To gNInc
        s = s & Chr(10) & gInc(i).Gravedad & " | " & gInc(i).Referencia & " | " & gInc(i).Texto
    Next i
    TestDiario2 = s
    Exit Function
EH:
    TestDiario2 = "ERR " & Err.Number & " " & Err.Description
End Function
'''.replace("__CSV__", str(CSV)).replace("__OUT__", str(OUT)).replace("__CSV2__", str(CSV2)).replace("__OUT2__", str(OUT2))
code = merge(["modA3Nucleo.bas", "modA3Config.bas", "modA3Diario.bas"], TEST)
(T / "salida" / "lo_merged_diario.bas").write_text(code, encoding="utf-8")
if OUT.exists(): OUT.unlink()
r = lo_runner.run([("Merged", code)], "Merged.TestDiario")
print(r)
if r is None: sys.exit("No compila en LibreOffice")
esperado = oraculo_diario.esperado()
obtenido = OUT.read_bytes().decode("ascii").split("\r\n")[:-1] if OUT.exists() else []
ok = obtenido == esperado and OUT.read_bytes().endswith(b"\r\n")
if not ok:
    for i, (a, b) in enumerate(zip(obtenido, esperado)):
        if a != b: print("DIF línea", i + 1, "\n obt:", a, "\n esp:", b); break
    print("nº líneas obtenido", len(obtenido), "esperado", len(esperado))
print("DIARIO OK: coincide con el oráculo" if ok else "DIARIO DIFERENTE")

# --- escenario 2: columnas por nombre, decimales con punto, asiento de una línea, descuadre incluido
if OUT2.exists(): OUT2.unlink()
r2 = lo_runner.run([("Merged", code)], "Merged.TestDiario2")
print(r2)
esp2 = oraculo_diario.esperado2()
obt2 = OUT2.read_bytes().decode("ascii").split("\r\n")[:-1] if OUT2.exists() else []
ok2 = obt2 == esp2 and r2 is not None and "Asiento de una sola línea" in r2 and "descuadrado" in r2
if not ok2:
    for i, (a, b) in enumerate(zip(obt2, esp2)):
        if a != b: print("DIF línea", i + 1, "\n obt:", a, "\n esp:", b); break
    print("nº líneas obtenido", len(obt2), "esperado", len(esp2))
print("DIARIO 2 OK" if ok2 else "DIARIO 2 DIFERENTE")
sys.exit(0 if (ok and ok2) else 1)
