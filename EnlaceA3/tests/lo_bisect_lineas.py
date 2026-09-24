"""Dentro del procedimiento que no compila, busca la línea culpable comentando
mitades de las líneas 'simples' (las que no abren ni cierran bloques)."""
import re, sys, pathlib
import lo_runner
code = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8").splitlines()
proc = sys.argv[2]
start = next(i for i,l in enumerate(code) if re.match(rf"^\s*(Public |Private )?(Function|Sub) {proc}\b", l))
end = next(i for i in range(start+1, len(code)) if re.match(r"^\s*End (Function|Sub)\s*$", code[i]))
BLOCK = re.compile(r"^\s*(If\b.*\bThen\s*$|ElseIf\b|Else\s*$|End If|Do\b|Loop\b|For\b|Next\b|With\b|End With|Select Case|Case\b|End Select|Dim\b)", re.I)
simple = [i for i in range(start+1, end) if code[i].strip() and not code[i].strip().startswith("'") and not BLOCK.match(code[i])
          and not code[i-1].rstrip().endswith(" _")]
# incluir las continuaciones de línea con su línea inicial
def group(i):
    g=[i]; j=i
    while code[j].rstrip().endswith(" _"): j+=1; g.append(j)
    return g
PROBE = "\nFunction ZZProbe() As String\n ZZProbe = \"ok\"\nEnd Function\n"
def compiles(commented):
    c = list(code)
    for i in commented:
        for j in group(i): c[j] = "'" + c[j]
    return lo_runner.run([("M", "\n".join(c) + PROBE)], "M.ZZProbe") == "ok"
cand = simple
if compiles([]): print("compila"); sys.exit()
if not compiles(cand): print("No es una línea simple (revisar líneas de bloque)"); sys.exit()
while len(cand) > 1:
    half = cand[:len(cand)//2]
    if compiles(half): cand = half
    else: cand = cand[len(cand)//2:]
for j in group(cand[0]): print(j+1, code[j])
