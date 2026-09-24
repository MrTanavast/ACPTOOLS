"""Une los módulos VBA de src/ en un solo módulo LibreOffice Basic (Option VBASupport)
para poder ejecutar el núcleo de la herramienta fuera de Excel.
LibreOffice no comparte tipos (Type) entre módulos, por eso se unen."""
import re, sys, pathlib
SRC = pathlib.Path(__file__).resolve().parent.parent / "src"
def merge(names, extra=""):
    out = ["Option VBASupport 1", "Option Explicit"]
    seen = {}
    for n in names:
        code = (SRC / n).read_text(encoding="utf-8")
        for line in code.splitlines():
            if re.match(r"^\s*(Attribute\s+VB_|Option\s+Explicit|VERSION\s|BEGIN|END$|\s*MultiUse)", line, re.I):
                continue
            m = re.match(r"^\s*(?:Public|Private)?\s*(?:Function|Sub)\s+(\w+)", line, re.I)
            if m:
                k = m.group(1).lower()
                if k in seen:
                    raise SystemExit(f"Procedimiento duplicado al unir: {m.group(1)} ({seen[k]} y {n})")
                seen[k] = n
            out.append(line)
    out.append(extra)
    return "\n".join(out)
if __name__ == "__main__":
    print(merge(sys.argv[1:]))
