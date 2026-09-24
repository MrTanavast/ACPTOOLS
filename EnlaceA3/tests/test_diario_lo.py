"""Prueba del motor de diarios en LibreOffice contra el oráculo Python."""
import pathlib, sys
from lo_merge import merge
import lo_runner, oraculo_diario
T = pathlib.Path(__file__).resolve().parent
CSV = T / "datos" / "diario_casos.csv"
OUT = T / "salida" / "SUENLACE_lo_diario.DAT"
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
'''.replace("__CSV__", str(CSV)).replace("__OUT__", str(OUT))
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
sys.exit(0 if ok else 1)
