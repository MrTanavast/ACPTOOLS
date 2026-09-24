"""Ejecuta todas las pruebas en LibreOffice."""
import subprocess, sys, pathlib
T = pathlib.Path(__file__).resolve().parent
fallos = 0
for t in sorted(T.glob("test_*_lo.py")):
    print(f"== {t.name}")
    r = subprocess.run([sys.executable, str(t)], cwd=T)
    fallos += r.returncode != 0
print("TODO OK" if fallos == 0 else f"{fallos} prueba(s) con fallos")
sys.exit(1 if fallos else 0)
