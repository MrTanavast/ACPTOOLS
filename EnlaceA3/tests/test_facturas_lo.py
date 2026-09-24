"""Prueba de regresión: ejecuta el motor VBA de facturas en LibreOffice con el CSV
real de Contapluz y compara el resultado byte a byte con el SUENLACE.DAT de referencia."""
import pathlib, sys, filecmp
from lo_merge import merge
import lo_runner
T = pathlib.Path(__file__).resolve().parent
CSV = T / "datos" / "JULIO_CONTAPLUZ.csv"
REF = T / "datos" / "SUENLACE_referencia_01692.DAT"
OUT = T / "salida" / "SUENLACE_lo_01692.DAT"
OUT.parent.mkdir(exist_ok=True)
TEST = r'''
Function TestFacturas() As String
    On Error GoTo EH
    Dim emp As TEmpresa, per As TPerfil, fil As TFiltros, res As TResumen, msg As String, datos As Variant, s As String, i As Long
    emp.Codigo = "01692": emp.Digitos = 8
    emp.CtaClientes = "43000000": emp.DescClientes = "Clientes varios Patio Posadero"
    emp.CtaVentas = "70500001": emp.DescVentas = "Prestacion servicios otros"
    emp.CtaVentasAlt = "70500000": emp.DescVentasAlt = "Prestacion servicios habitacion"
    emp.CtaRetencion = "47300000"
    If Not PerfilPredefinido("CONTAPLUZ", per) Then TestFacturas = "ERR perfil": Exit Function
    IvaReiniciar
    IvaAgregar 4, "47700004", 0, ""
    IvaAgregar 10, "47700010", 0, ""
    IvaAgregar 21, "47700021", 0, ""
    fil.InclFacturas = True: fil.InclTickets = True: fil.InclAbonos = True
    datos = ParsearCSV(LeerFicheroTexto("__CSV__"), per.Separador)
    If Not FacturasProcesar(emp, per, fil, datos, res, msg) Then TestFacturas = "ERR " & msg: Exit Function
    EscribirDat "__OUT__"
    s = "docs=" & res.UnidadesLeidas & " exp=" & res.UnidadesExportadas & " excl=" & res.UnidadesExcluidas & _
        " lineas=" & gNDat & " fact=" & res.NumFacturas & " tick=" & res.NumTickets & " abon=" & res.NumAbonos & _
        " base=" & ImporteTexto(res.BaseImp) & " cuota=" & ImporteTexto(res.Cuota) & " total=" & ImporteTexto(res.Total) & _
        " origen=" & ImporteTexto(res.TotalOrigen) & " excluido=" & ImporteTexto(res.TotalExcluido) & " ajuste=" & ImporteTexto(res.AjusteSignoAbonos)
    For i = 1 To gNInc
        s = s & Chr(10) & gInc(i).Gravedad & " | " & gInc(i).Referencia & " | fila " & gInc(i).Fila & " | " & gInc(i).Texto & " | " & gInc(i).Tratamiento
    Next i
    TestFacturas = s
    Exit Function
EH:
    TestFacturas = "ERR " & Err.Number & " " & Err.Description & " (" & Err.Source & ")"
End Function
'''.replace("__CSV__", str(CSV)).replace("__OUT__", str(OUT))
code = merge(["modA3Nucleo.bas", "modA3Config.bas", "modA3Facturas.bas"], TEST)
(T / "salida" / "lo_merged_facturas.bas").write_text(code, encoding="utf-8")
if OUT.exists(): OUT.unlink()
r = lo_runner.run([("Merged", code)], "Merged.TestFacturas")
print(r)
if r is None: sys.exit("La macro no se ha ejecutado (error de compilación en LibreOffice)")
ok = OUT.exists() and OUT.read_bytes() == REF.read_bytes()
print("IDENTICO AL DAT DE REFERENCIA" if ok else "DIFERENTE del DAT de referencia")
sys.exit(0 if ok else 1)
