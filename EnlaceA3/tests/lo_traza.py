"""Instrumenta el código unido: antes de cada sentencia simple de cada procedimiento
guarda el número de línea ORIGINAL en gZZTr, para saber dónde salta un error."""
import re
BLOCKSTART = re.compile(r"^\s*(ElseIf\b|Else\s*$|End If|Loop\b|Next\b|End With|Case\b|End Select|End Sub|End Function|End Type)", re.I)
PROC = re.compile(r"^\s*(Public |Private )?(Function|Sub) (\w+)", re.I)
def instrumentar_todo(code):
    lines = code.splitlines(); out = []; dentro = False; nombre = ""
    for i, l in enumerate(lines):
        m = PROC.match(l)
        if m: dentro = True; nombre = m.group(3); out.append(l); continue
        if re.match(r"^\s*End (Function|Sub)\s*$", l): dentro = False; out.append(l); continue
        prev = lines[i-1].rstrip() if i else ""
        if dentro and not nombre.startswith(("Test", "ZZ")) and l.strip() and not l.strip().startswith("'") \
           and not prev.endswith(" _") and not BLOCKSTART.match(l) and not re.match(r"^\s*Dim\b", l) \
           and not re.match(r"^\s*\w+:\s*$", l) and not prev.endswith(" _"):
            out.append(f"gZZTr = {i+1}")
        out.append(l)
    return "\n".join(out)
def ejecutar_con_traza(code, entrada):
    import lo_runner
    inst = instrumentar_todo(code).replace("Option VBASupport 1\nOption Explicit", "Option VBASupport 1\nOption Explicit\nPublic gZZTr As Long", 1)
    return lo_runner.run([("M", inst)], entrada), code.splitlines()
def informar(r, src):
    print((r or "None")[:400])
    m = re.search(r"linea (\d+)", r or "")
    if m:
        n = int(m.group(1)); print(n, src[n-1])
        for k in range(n-1, 0, -1):
            if PROC.match(src[k]): print("en:", src[k].strip()); break
