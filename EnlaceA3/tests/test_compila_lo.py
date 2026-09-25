"""Comprueba que TODO el código compila (sintaxis) en LibreOffice: los módulos estándar
y las dos clases (adaptadas: MSForms -> Object, Me -> objeto auxiliar, New cls -> función).
No sustituye a 'Depuración > Compilar' de Excel, pero detecta líneas partidas, bucles mal
cerrados, sentencias incompletas, etc."""
import re, sys, pathlib
import lo_runner
from lo_merge import merge
SRC = pathlib.Path(__file__).resolve().parent.parent / "src"
MODULOS = ["modA3Nucleo.bas", "modA3Config.bas", "modA3Facturas.bas", "modA3Diario.bas",
           "modA3Control.bas", "modA3Inicio.bas"]
PROBE = "\nFunction ZZProbe() As String\n ZZProbe = \"ok\"\nEnd Function\n"
AUX = "\nPrivate ZZYo As Object\nFunction ZZNuevo() As Object\nEnd Function\n"

def adaptar(texto):
    t = re.sub(r"^\s*Attribute\s+\w+\.VB_\w+.*$", "", texto, flags=re.M)
    t = re.sub(r"Public WithEvents (\w+) As MSForms\.\w+", r"Public \1 As Object", t)
    t = re.sub(r"As MSForms\.\w+", "As Object", t)
    t = re.sub(r"New clsA3\w+", "ZZNuevo()", t)
    t = re.sub(r"As clsA3\w+", "As Object", t)
    t = re.sub(r"\bMe\b", "ZZYo", t)
    return t

def compila(extra, nombre):
    # sustituir referencias a las clases también en los módulos estándar
    codigo = adaptar(merge(MODULOS, extra + AUX + PROBE))
    r = lo_runner.run([("M", codigo)], "M.ZZProbe")
    print(f"{nombre}: {'OK' if r == 'ok' else 'NO COMPILA'}")
    return r == "ok"

ok = compila("", "módulos estándar")
for cls in ["clsA3Ventana.cls", "clsA3Evento.cls"]:
    texto = (SRC / cls).read_text(encoding="utf-8")
    texto = re.sub(r"^(VERSION .*|BEGIN|END|\s*MultiUse.*|Attribute VB_.*|Option Explicit)\s*$", "", texto, flags=re.M)
    ok = compila(texto, cls) and ok
print("COMPILA TODO" if ok else "HAY ERRORES DE COMPILACIÓN")
sys.exit(0 if ok else 1)
