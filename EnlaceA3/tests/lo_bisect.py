"""Localiza el procedimiento que LibreOffice no compila (búsqueda binaria)."""
import re, sys, pathlib
import lo_runner
code = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8").splitlines()
# trocear en bloques: cabecera + procedimientos
starts = [i for i,l in enumerate(code) if re.match(r"^\s*(Public |Private )?(Function|Sub) \w+", l)]
head = code[:starts[0]]
blocks = []
for k,s in enumerate(starts):
    e = starts[k+1] if k+1 < len(starts) else len(code)
    blocks.append(code[s:e])
PROBE = "\nFunction ZZProbe() As String\n ZZProbe = \"ok\"\nEnd Function\n"
def compiles(n, headlines=None):
    body = "\n".join((headlines or head) + [l for b in blocks[:n] for l in b]) + PROBE
    return lo_runner.run([("M", body)], "M.ZZProbe") == "ok"
if not compiles(0):
    # problema en la cabecera: bisect líneas de cabecera
    units=[]; cur=[]
    for l in head:
        cur.append(l)
        if re.match(r"^\s*(Public |Private )?Type ", l): continue
        if cur and re.match(r"^\s*(Public |Private )?Type ", cur[0]) and not re.match(r"^\s*End Type", l): continue
        units.append(cur); cur=[]
    if cur: units.append(cur)
    lo, hi = 0, len(units)
    while hi - lo > 1:
        mid = (lo+hi)//2
        if compiles(0, [l for u in units[:mid] for l in u]): lo = mid
        else: hi = mid
    print("Cabecera falla en:", "\n".join(units[hi-1])); sys.exit()
lo, hi = 0, len(blocks)
if compiles(hi): print("Todo compila"); sys.exit()
while hi - lo > 1:
    mid = (lo+hi)//2
    if compiles(mid): lo = mid
    else: hi = mid
print("Falla el bloque", hi, ":\n" + "\n".join(blocks[hi-1][:60]))
