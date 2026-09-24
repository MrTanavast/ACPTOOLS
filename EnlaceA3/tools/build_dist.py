"""Genera la carpeta dist/ lista para importar en Excel:
   - .bas / .cls en Windows-1252 y con saltos de línea CR/LF (el editor de VBA
     no lee bien UTF-8 ni saltos LF: las tildes saldrían mal).
   - el código del formulario como .txt para pegar.
   Uso:  python3 EnlaceA3/tools/build_dist.py"""
import pathlib, shutil, sys
RAIZ = pathlib.Path(__file__).resolve().parent.parent
SRC, DIST = RAIZ / "src", RAIZ / "dist"
DIST.mkdir(exist_ok=True)
MODULOS = ["modA3Nucleo.bas", "modA3Config.bas", "modA3Facturas.bas", "modA3Diario.bas",
           "modA3Control.bas", "modA3Inicio.bas", "clsA3Evento.cls"]
def convertir(origen, destino):
    texto = origen.read_text(encoding="utf-8").replace("\r\n", "\n")
    try:
        datos = texto.replace("\n", "\r\n").encode("cp1252")
    except UnicodeEncodeError as e:
        sys.exit(f"{origen.name}: carácter no representable en Windows-1252: {texto[e.start:e.end]!r} (usa ChrW)")
    destino.write_bytes(datos)
for m in MODULOS:
    convertir(SRC / m, DIST / m)
convertir(SRC / "frmEnlaceA3.codigo.bas", DIST / "frmEnlaceA3_codigo.txt")
for extra in ["logo_acp.png"]:
    if (RAIZ / "recursos" / extra).exists():
        shutil.copy(RAIZ / "recursos" / extra, DIST / extra)
print("dist/ generado:", ", ".join(sorted(p.name for p in DIST.iterdir())))
