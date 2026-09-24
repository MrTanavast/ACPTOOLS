"""Oráculo independiente (Python) del libro de FACTURAS EMITIDAS, perfil GENERAL.

Genera los registros tipo 1/2 (cabecera) y 9 (detalle de IVA) del SUENLACE.DAT
(formato 4: 254 posiciones) a partir de la especificación de a3 (posiciones de cada
campo) y de las reglas documentadas del procedimiento (cabecera de modA3Facturas.bas,
instrucciones de PLANTILLA_FACTURAS y ayuda del formulario). No reutiliza nada del VBA.

Donde la regla es de negocio y no de formato, se implementa lo que DEBERÍA pasar:
  - Total vacío en una fila = base + cuota + recargo - retención (PLANTILLA_FACTURAS: U5).
  - El filtro de tipos se aplica al tipo con el que se exporta el documento (un documento
    sin tipo y con total negativo es un abono).
  - Una serie es el prefijo del documento hasta un límite natural: en "F-2026-001" las
    series "F" y "F-2026" (el filtro por serie acepta cualquiera de las dos).
  - Un % (IVA, recargo o retención) leído de una celda numérica entre 0 y 1 es un
    porcentaje con formato % de Excel (0,052 -> 5,20).
  - Entregas intracomunitarias (subtipo 03/04) acumulan al impreso 349 (02), no al 347 (01).
"""
import datetime
import re
import unicodedata
from decimal import Decimal, ROUND_HALF_UP

LONG = 254
CERO = Decimal("0.00")

# --------------------------------------------------------------------------------
#  Formato de campos (especificación a3)
# --------------------------------------------------------------------------------
_EXTRA = {"–": "-", "—": "-", "€": "EUR", "º": "o", "ª": "a"}


def ascii_a3(s):
    s = "".join(_EXTRA.get(ch, ch) for ch in str(s))
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = "".join(ch if 32 <= ord(ch) <= 126 else " " for ch in s)
    return " ".join(s.split())


def txt(s, n):
    return ascii_a3(s)[:n].ljust(n)


def imp(x):
    x = Decimal(x).quantize(CERO)
    assert abs(x) < Decimal("1E10")
    return ("-" if x < 0 else "+") + f"{abs(x):013.2f}"


def pct(x):
    x = abs(Decimal(x)).quantize(CERO)
    return f"{x:05.2f}"


def f8(d):
    return d.strftime("%Y%m%d")


def registro(campos):
    """campos: lista de (posición inicial, posición final, valor) con posiciones de la
    especificación (1 = primera). Lo no informado queda en blanco."""
    buf = [" "] * LONG
    usados = set()
    for ini, fin, val in campos:
        assert len(val) == fin - ini + 1, (ini, fin, val)
        for p in range(ini, fin + 1):
            assert p not in usados, f"posición {p} solapada"
            usados.add(p)
        buf[ini - 1:fin] = list(val)
    return "".join(buf)


def reg_cabecera(emp, fecha, abono, cuenta, desc_cuenta, doc, desc, total, nif, nombre, cp, fecha_op, fecha_fac):
    return registro([
        (1, 1, "4"), (2, 6, emp), (7, 14, f8(fecha)), (15, 15, "2" if abono else "1"),
        (16, 27, txt(cuenta, 12)), (28, 57, txt(desc_cuenta, 30)),
        (58, 58, "1"),                                    # tipo de factura: 1 = ventas
        (59, 68, txt(doc, 10)), (69, 69, "I"), (70, 99, txt(desc, 30)), (100, 113, imp(total)),
        # 114-175 reserva
        (176, 189, txt(nif, 14)), (190, 229, txt(nombre, 40)), (230, 234, txt(cp, 5)),
        # 235-236 reserva
        (237, 244, f8(fecha_op)), (245, 252, f8(fecha_fac)),
        (253, 253, "E"), (254, 254, "N")])


def reg_detalle(emp, fecha, cuenta, desc_cuenta, doc, imu, desc, subtipo, base, piva, cuota,
                pre, cre, pret, cret, impreso, cta_iva, cta_re, cta_ret):
    return registro([
        (1, 1, "4"), (2, 6, emp), (7, 14, f8(fecha)), (15, 15, "9"),
        (16, 27, txt(cuenta, 12)), (28, 57, txt(desc_cuenta, 30)),
        (58, 58, "C"),                                    # cargo en factura
        (59, 68, txt(doc, 10)), (69, 69, imu), (70, 99, txt(desc, 30)),
        (100, 101, subtipo), (102, 115, imp(base)), (116, 120, pct(piva)), (121, 134, imp(cuota)),
        (135, 139, pct(pre)), (140, 153, imp(cre)), (154, 158, pct(pret)), (159, 172, imp(cret)),
        (173, 174, impreso), (175, 175, "S"), (176, 176, "N"),
        # 177 criterio de caja, 178 marca IVA 0 % (blanco = exento), 179-191 reserva
        (192, 203, txt(cta_iva, 12)), (204, 215, txt(cta_re, 12)), (216, 227, txt(cta_ret, 12)),
        # 228-251 cuentas "2" (autorrepercusión) en blanco, 252 analítica
        (253, 253, "E"), (254, 254, "N")])


# --------------------------------------------------------------------------------
#  Lectura de valores
# --------------------------------------------------------------------------------
def q2(x):
    return Decimal(x).quantize(CERO, rounding=ROUND_HALF_UP)


def vacio(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def parse_importe(v):
    """Devuelve (Decimal con 2 decimales, ok)."""
    if vacio(v):
        return CERO, True
    if isinstance(v, (int, float, Decimal)):
        return q2(Decimal(repr(v)) if isinstance(v, float) else Decimal(v)), True
    s = str(v).replace(" ", "").replace("€", "").replace(" ", "")
    s = re.sub("EUR", "", s, flags=re.I)
    if s == "":
        return CERO, True
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    if s.endswith("-"):
        neg, s = not neg, s[:-1]
    if s.startswith("-"):
        neg, s = not neg, s[1:]
    elif s.startswith("+"):
        s = s[1:]
    if not s or not re.fullmatch(r"[0-9.,]+", s):
        return None, False
    if "." in s and "," in s:
        dec = "." if s.rfind(".") > s.rfind(",") else ","
    elif "," in s:
        dec = "," if s.count(",") == 1 else "."
    elif s.count(".") == 1:
        dec = "," if (len(s) - s.rfind(".") - 1 == 3 and not s.startswith("0")) else "."
    else:
        dec = ","
    miles = "," if dec == "." else "."
    if s.count(dec) > 1:
        return None, False
    s = s.replace(miles, "").replace(dec, ".")
    if s in ("", "."):
        return None, False
    x = q2(Decimal(s))
    return (-x if neg else x), True


def parse_pct(v):
    """% de IVA, recargo o retención. Celda numérica entre 0 y 1 = formato % de Excel."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 0 < v < 1:
        return q2(Decimal(repr(v)) * 100), True
    return parse_importe(v)


def parse_fecha(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = texto(v).split(" ")[0]
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        a, me, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})", s)
        if not m:
            return None
        d, me, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a < 100:
            a += 2000
    try:
        return datetime.date(a, me, d)
    except ValueError:
        return None


def norm(s):
    s = ascii_a3(str(s).replace("%", " PCT ")).upper()
    return " ".join(re.sub(r"[^A-Z0-9]", " ", s).split())


def norm_cuenta(c, dig):
    c = texto(c).replace(" ", "")
    if not c:
        return None
    if "." in c:
        if c.count(".") > 1:
            return None
        izq, der = c.split(".")
        if not izq or len(izq) + len(der) > dig:
            return None
        c = izq + "0" * (dig - len(izq) - len(der)) + der
    if not c.isdigit() or len(c) != dig:
        return None
    return c


def serie_y_numero(doc):
    m = re.fullmatch(r"(.*?)(\d*)", texto(doc))
    return norm(m.group(1)), (int(m.group(2)) if m.group(2) else None)


def series_candidatas(doc):
    """Prefijos de la serie que una persona escribiría: "F-2026-001" -> {"F", "F 2026"}."""
    s, _ = serie_y_numero(doc)
    cand = {s}
    for i in range(1, len(s)):
        a, b = s[i - 1], s[i]
        if b == " " or (a.isalpha() and b.isdigit()) or (a.isdigit() and b.isalpha()):
            cand.add(s[:i].strip())
    cand.discard("")
    return cand


# --------------------------------------------------------------------------------
#  Motor del oráculo (perfil GENERAL)
# --------------------------------------------------------------------------------
ABONO = {"ABONO", "A", "RECTIFICATIVA", "R", "NOTA DE CREDITO", "NC"}
TICKET = {"TICKET", "T", "SIMPLIFICADA", "S", "FACTURA SIMPLIFICADA"}
NO_ANULADA = {"NO", "N", "0", "FALSO", "FALSE"}
NOMBRE_VACIO = "Clientes varios"


class Empresa:
    def __init__(self, codigo, digitos, cta_clientes, desc_clientes, cta_ventas, desc_ventas, cta_ret, iva):
        self.codigo, self.digitos = codigo, digitos
        self.cta_clientes, self.desc_clientes = cta_clientes, desc_clientes
        self.cta_ventas, self.desc_ventas = cta_ventas, desc_ventas
        self.cta_ret = cta_ret
        self.iva = iva      # {Decimal pct: (cuenta, Decimal pct_re, cuenta_re)}


def leer_csv(ruta):
    raw = open(ruta, "rb").read()
    try:
        t = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        t = raw.decode("cp1252")
    filas = [l.split(";") for l in t.replace("\r\n", "\n").split("\n") if l.strip()]
    return filas


def _col(cab, nombre):
    for i, c in enumerate(cab):
        if norm(c) == norm(nombre):
            return i
    return None


def leer_lineas(filas, emp):
    cab = filas[0]
    C = {k: _col(cab, n) for k, n in [
        ("fecha", "Fecha"), ("doc", "Nº Factura"), ("tipo", "Tipo"), ("cliente", "Cliente"), ("nif", "NIF"),
        ("cp", "CP"), ("base", "Base imponible"), ("piva", "% IVA"), ("cuota", "Cuota IVA"), ("total", "Total"),
        ("cta_cli", "Cuenta cliente"), ("cta_ven", "Cuenta ventas"), ("nom_ven", "Nombre cuenta ventas"),
        ("pre", "% Recargo"), ("cre", "Cuota recargo"), ("pret", "% Retención"), ("cret", "Cuota retención"),
        ("subtipo", "Subtipo"), ("anulada", "Anulada")]}

    def g(fila, k):
        i = C[k]
        return None if i is None or i >= len(fila) else fila[i]

    L = []
    for n, fila in enumerate(filas[1:], start=2):
        if all(vacio(x) for x in fila):
            continue
        e = []
        d = {"fila": n}
        d["fecha"] = parse_fecha(g(fila, "fecha"))
        if d["fecha"] is None:
            e.append("fecha")
        d["doc"] = texto(g(fila, "doc"))
        if not d["doc"]:
            e.append("doc")
        d["tipo_txt"] = texto(g(fila, "tipo"))
        t = norm(d["tipo_txt"])
        d["tipo"] = "abono" if t in ABONO else ("ticket" if t in TICKET else "factura")
        a = texto(g(fila, "anulada"))
        d["anulada"] = bool(a) and norm(a) not in NO_ANULADA
        d["cliente"], d["nif"], d["cp"] = texto(g(fila, "cliente")), texto(g(fila, "nif")), texto(g(fila, "cp"))
        for k, parser in [("base", parse_importe), ("piva", parse_pct), ("cuota", parse_importe),
                          ("pre", parse_pct), ("cre", parse_importe), ("pret", parse_pct), ("cret", parse_importe)]:
            d[k], ok = parser(g(fila, k))
            if not ok:
                e.append(k)
        if C["total"] is not None and not vacio(g(fila, "total")):
            d["total"], ok = parse_importe(g(fila, "total"))
            if not ok:
                e.append("total")
        elif not e:
            d["total"] = d["base"] + d["cuota"] + d["cre"] - d["cret"]
        else:
            d["total"] = CERO
        # cuentas
        cc = texto(g(fila, "cta_cli"))
        if cc:
            d["cta_cli"], d["desc_cli"] = cc, (d["cliente"] or "Cliente " + cc)
        else:
            d["cta_cli"], d["desc_cli"] = emp.cta_clientes, emp.desc_clientes
        cv = texto(g(fila, "cta_ven"))
        if cv:
            d["cta_ven"] = cv
            d["desc_ven"] = texto(g(fila, "nom_ven")) or (emp.desc_ventas if cv == emp.cta_ventas else "Ventas")
        else:
            d["cta_ven"], d["desc_ven"] = emp.cta_ventas, emp.desc_ventas
        cfg = emp.iva.get(d["piva"]) if d["piva"] is not None else None
        d["cta_iva"] = cfg[0] if cfg else ""
        d["cta_re"] = cfg[2] if cfg else ""
        if d["cre"] and d["pre"] == 0 and cfg:
            d["pre"] = cfg[1]
        d["cta_ret"] = emp.cta_ret
        st = texto(g(fila, "subtipo"))
        if st == "":
            d["subtipo"] = "01"
        elif st.isdigit() and len(st) <= 2 and int(st) > 0:
            d["subtipo"] = st.zfill(2)
        else:
            d["subtipo"] = "??"
            e.append("subtipo")
        d["errores"] = e
        L.append(d)
    return L


def procesar(filas, emp, fil):
    """fil: dict con desde, hasta (date), ref_desde, ref_hasta, incl, excl (str),
    facturas, tickets, abonos (bool). Devuelve dict con lineas, docs (orden), excluidos,
    avisos y resumen."""
    L = leer_lineas(filas, emp)
    # 1) filtros por línea (fecha, rango, series)
    dentro = []
    for d in L:
        ok = True
        if d["fecha"] is not None:
            if fil.get("desde") and d["fecha"] < fil["desde"]:
                ok = False
            if fil.get("hasta") and d["fecha"] > fil["hasta"]:
                ok = False
        rd, rh = fil.get("ref_desde", ""), fil.get("ref_hasta", "")
        if ok and (rd or rh):
            ref = serie_y_numero(rd or rh)[0]
            s, num = serie_y_numero(d["doc"])
            if s != ref or num is None:
                ok = False
            if ok and rd and serie_y_numero(rd)[1] is not None and num < serie_y_numero(rd)[1]:
                ok = False
            if ok and rh and serie_y_numero(rh)[1] is not None and num > serie_y_numero(rh)[1]:
                ok = False
        if ok and fil.get("incl"):
            ok = bool(series_candidatas(d["doc"]) & {norm(x) for x in fil["incl"].split(";") if norm(x)})
        if ok and fil.get("excl"):
            ok = not (series_candidatas(d["doc"]) & {norm(x) for x in fil["excl"].split(";") if norm(x)})
        dentro.append(ok)
    # 2) agrupación por documento + fecha
    docs, idx = [], {}
    for d, ok in zip(L, dentro):
        if not ok:
            continue
        k = (d["doc"].upper(), d["fecha"] if d["fecha"] else "?" + d["doc"])
        if k not in idx:
            idx[k] = len(docs)
            docs.append([])
        docs[idx[k]].append(d)

    # 3) filtro de tipos: sobre el tipo con el que se exporta el documento
    def tipo_doc(ls):
        if ls[0]["tipo"] == "abono":
            return "abono"
        if all(x["tipo_txt"] == "" for x in ls) and sum(x["total"] for x in ls) < 0:
            return "abono"
        return ls[0]["tipo"]
    quedan = []
    filtradas = sum(1 for x in dentro if not x)
    for ls in docs:
        t = tipo_doc(ls)
        if {"factura": fil.get("facturas", True), "ticket": fil.get("tickets", True),
                "abono": fil.get("abonos", True)}[t]:
            quedan.append(ls)
        else:
            filtradas += len(ls)

    R = {"lineas": [], "excluidos": {}, "avisos": set(), "exportados": [],
         "res": dict(leidas=len(quedan), exp=0, excl=0, filtradas=filtradas, fact=0, tick=0, abon=0,
                     base=CERO, cuota=CERO, cuota_re=CERO, cuota_ret=CERO, total=CERO)}
    for ls in quedan:
        _documento(ls, emp, tipo_doc(ls), R)
    R["res"]["lineas"] = len(R["lineas"])
    return R


def _documento(ls, emp, tipo, R):
    doc, fecha = ls[0]["doc"], ls[0]["fecha"]

    def excluir(cat):
        R["excluidos"][doc] = cat
        R["res"]["excl"] += 1

    if any(x["anulada"] for x in ls):
        return excluir("ANULADO")
    if any(x["errores"] for x in ls):
        return excluir("DATOS")
    if sum(abs(x["total"]) for x in ls) == 0:
        return excluir("CERO")
    vivas = [x for x in ls if not (x["base"] == 0 and x["cuota"] == 0 and x["cre"] == 0 and x["cret"] == 0)]
    cero_con_total = sum(x["total"] for x in ls if x not in vivas)
    if not vivas:
        return excluir("SINLINEAS")
    for x in vivas:
        if x["piva"] < 0 or x["piva"] >= 100 or (x["piva"] == 0 and x["cuota"] != 0) or not x["cta_iva"]:
            return excluir("IVA")
    cuentas = []
    for x in vivas:
        c = {"cli": norm_cuenta(x["cta_cli"], emp.digitos), "ven": norm_cuenta(x["cta_ven"], emp.digitos),
             "iva": norm_cuenta(x["cta_iva"], emp.digitos),
             "re": norm_cuenta(x["cta_re"], emp.digitos) if x["cre"] else "",
             "ret": norm_cuenta(x["cta_ret"], emp.digitos) if x["cret"] else ""}
        if None in c.values():
            return excluir("CUENTA")
        cuentas.append(c)
    total = sum(x["total"] for x in vivas)
    if total != sum(x["base"] + x["cuota"] + x["cre"] - x["cret"] for x in vivas):
        return excluir("DESCUADRE")

    abono = tipo == "abono"
    factor = -1 if (abono and total < 0) else 1
    if abono and total > 0:
        R["avisos"].add((doc, "ABONO_POSITIVO"))
    if not abono and total < 0:
        R["avisos"].add((doc, "NEGATIVA"))
    if cero_con_total != 0:
        R["avisos"].add((doc, "CERO_CON_TOTAL"))
    if len(doc) > 10:
        R["avisos"].add((doc, "DOC_LARGO"))
    for x in vivas:
        if x["piva"] > 0:
            esperada = q2(x["base"] * x["piva"] / 100)
            if abs(esperada - x["cuota"]) > Decimal("0.02") + abs(x["base"]) * Decimal("0.0005"):
                R["avisos"].add((doc, "CUOTA"))

    p = ls[0]
    cliente = ascii_a3(p["cliente"]) or NOMBRE_VACIO
    nif = ascii_a3(p["nif"])
    cp = p["cp"] if nif else ""
    desc = doc + " " + cliente
    lineas = [reg_cabecera(emp.codigo, fecha, abono, cuentas[0]["cli"], p["desc_cli"], doc, desc, factor * total,
                           nif, cliente if nif else "", cp, fecha, fecha)]
    # detalle: una línea por cuenta de ventas + tipo de IVA (+ recargo, retención, subtipo)
    det = {}
    for x, c in zip(vivas, cuentas):
        k = (c["ven"], x["desc_ven"], x["piva"], c["iva"], x["pre"], c["re"], x["pret"], c["ret"], x["subtipo"])
        acc = det.setdefault(k, [CERO, CERO, CERO, CERO])
        for j, campo in enumerate(("base", "cuota", "cre", "cret")):
            acc[j] += x[campo]
    claves = sorted(det, key=lambda k: (k[0], k[1], k[2], k[4], k[6]))
    signo = -1 if abono else 1
    for n, k in enumerate(claves):
        ven, dven, piva, civa, pre, cre_cta, pret, cret_cta, st = k
        b, cu, re_, rt = (factor * v for v in det[k])
        impreso = "02" if st in ("03", "04") else "01"
        lineas.append(reg_detalle(emp.codigo, fecha, ven, dven, doc, "U" if n == len(claves) - 1 else "M", desc,
                                  st, b, piva, cu, pre, re_, pret, rt, impreso, civa, cre_cta, cret_cta))
        r = R["res"]
        r["base"] += signo * b
        r["cuota"] += signo * cu
        r["cuota_re"] += signo * re_
        r["cuota_ret"] += signo * rt
        r["total"] += signo * (b + cu + re_ - rt)
    R["lineas"].extend(lineas)
    R["exportados"].append((doc, fecha, lineas))
    R["res"]["exp"] += 1
    R["res"]["abon" if abono else ("tick" if p["tipo"] == "ticket" else "fact")] += 1
