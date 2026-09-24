"""Oráculo independiente (Python) del libro de FACTURAS EMITIDAS, perfil GENERAL.

Genera los registros tipo 1/2 (cabecera) y 9 (detalle de IVA) del SUENLACE.DAT
(formato 4: 254 posiciones) a partir de la especificación de a3 (posiciones de cada
campo) y de las reglas documentadas del procedimiento (cabecera de modA3Facturas.bas,
instrucciones de PLANTILLA_FACTURAS y ayuda del formulario). No reutiliza nada del VBA.

Reglas de negocio que implementa (lo que DEBE pasar):
  - Total vacío en una fila = base + cuota + recargo - retención.
  - Tipo del documento = el de su primera línea. Una factura sin texto en TIPO en
    ninguna línea y cuyas líneas con importe suman un total negativo es un abono.
    El filtro de tipos se aplica al documento: tickets -> Tickets; abonos y
    rectificativas -> Abonos; el resto -> Facturas. Un documento fuera por tipo cuenta
    todas sus filas como filtradas.
  - % con formato % de Excel (celda numérica): IVA y retención entre 0 y 1, recargo
    entre 0 y 0,1 -> x 100 antes de redondear (0,052 -> 5,20; 0,5 de recargo es 0,50 %).
  - Series: el valor escrito vale si la serie normalizada es igual o empieza por él
    seguido de espacio, o de dígito cuando el valor no acaba en dígito
    ("F" -> "F 2026", "FV" -> "FV2026", pero "F" no -> "FV2026" y "F 20" no -> "F 2026").
  - Impreso: columna IMPRESO; si no, 02 (349) para subtipo 03/04 y 01 (347) en el resto.
  - Nº de factura en a3 (10 posiciones): ASCII; si es más largo, sin - / espacio . y,
    si aún no cabe, los 10 últimos. Si otro documento ya exportado llega a a3 con el
    mismo nº y fecha, el segundo se excluye.
  - Rectificativa: registro 2 con el signo del origen invertido siempre (una
    rectificativa positiva aumenta la factura: sale en negativo y con aviso).
    Abono: registro 2; en negativo en origen -> positivo; en positivo -> aviso y tal cual.
  - Retención con signo contrario a la base -> se invierte (aviso). Retención sin % ->
    cuota / base; si es 1, 2, 7, 15, 19 o 24 se usa (aviso), si no, error de lectura.
    Recargo con cuota y sin % (ni en la línea ni en la tabla de IVA) -> error de lectura.
  - Cuotas de IVA, recargo y retención comparadas con base x % (tolerancia
    0,02 + |base| x 0,0005) -> aviso si no cuadran.
  - Subtipo vacío -> 01 (aviso si el IVA es 0 % con base); válidos 01-06, 08 y 09.
  - NIF normalizado (mayúsculas, sin separadores, sin prefijo ES si el resto es válido);
    si no es un DNI / NIE / CIF español válido -> aviso. CP solo con NIF; una celda
    numérica de 4 dígitos lleva el 0 delante; si no son 5 dígitos -> en blanco y aviso.
  - Las líneas anuladas forman su propio documento (excluido); el documento no anulado
    con el mismo nº y fecha se exporta con aviso.
  - Una celda con error (#N/D...) en cualquier columna del perfil -> error de lectura.
"""
import datetime
import re
import unicodedata
from decimal import Decimal, ROUND_HALF_UP

LONG = 254
CERO = Decimal("0.00")


class ErrorCelda:
    """Celda de Excel con error de fórmula (#N/D, #¡VALOR!...)."""

    def __init__(self, texto="#N/D", codigo=2042):
        self.texto, self.codigo = texto, codigo

    def __repr__(self):
        return self.texto


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


def documento_a3(doc):
    """Nº de factura tal como cabe en a3 (posiciones 59-68)."""
    s = ascii_a3(doc)
    if len(s) > 10:
        s = re.sub(r"[-/ .]", "", s)
    return s[-10:]


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
        (59, 68, txt(documento_a3(doc), 10)), (69, 69, "I"), (70, 99, txt(desc, 30)), (100, 113, imp(total)),
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
        (59, 68, txt(documento_a3(doc), 10)), (69, 69, imu), (70, 99, txt(desc, 30)),
        (100, 101, subtipo), (102, 115, imp(base)), (116, 120, pct(piva)), (121, 134, imp(cuota)),
        (135, 139, pct(pre)), (140, 153, imp(cre)), (154, 158, pct(pret)), (159, 172, imp(cret)),
        (173, 174, impreso), (175, 175, "S"), (176, 176, "N"),
        # 177 criterio de caja y 178 marca IVA 0 % en blanco; 179-191 reserva
        (192, 203, txt(cta_iva, 12)), (204, 215, txt(cta_re, 12)), (216, 227, txt(cta_ret, 12)),
        # 228-251 cuentas de IVA / recargo "2" (autorrepercusión) en blanco, 252 analítica
        (253, 253, "E"), (254, 254, "N")])


# --------------------------------------------------------------------------------
#  Lectura de valores
# --------------------------------------------------------------------------------
def q2(x):
    return Decimal(x).quantize(CERO, rounding=ROUND_HALF_UP)


def vacio(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def es_numero(v):
    return isinstance(v, (int, float, Decimal)) and not isinstance(v, bool)


def texto(v):
    if v is None or isinstance(v, ErrorCelda):
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def parse_importe(v):
    """Devuelve (Decimal con 2 decimales, ok). Si no es numérico: (0, False)."""
    if isinstance(v, ErrorCelda):
        return CERO, False
    if vacio(v):
        return CERO, True
    if es_numero(v):
        return q2(Decimal(repr(v)) if isinstance(v, float) else Decimal(v)), True
    s = str(v).replace(" ", "").replace("€", "").replace(" ", "")
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
        return CERO, False
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
        return CERO, False
    s = s.replace(miles, "").replace(dec, ".")
    if s in ("", "."):
        return CERO, False
    x = q2(Decimal(s))
    if abs(x) >= Decimal("1E10"):
        return CERO, False
    return (-x if neg else x), True


def parse_pct(v, limite):
    """% de IVA, recargo o retención. Una celda numérica entre 0 y "limite" es un %
    con formato de Excel (0,21 = 21 %): se multiplica por 100 antes de redondear."""
    if es_numero(v) and 0 < v < limite:
        return q2(Decimal(repr(v)) * 100), True
    if isinstance(v, str) and v.strip().endswith("%"):
        v = v.strip()[:-1]
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
    if not 1990 <= a <= 2099:
        return None
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
    if not re.fullmatch(r"[0-9]+", c) or len(c) != dig:
        return None
    return c


# --------------------------------------------------------------------------------
#  NIF español (DNI, NIE, CIF)
# --------------------------------------------------------------------------------
LETRAS_DNI = "TRWAGMYFPDXBNJZSQVHLCKE"


def control_cif(cuerpo7):
    """Carácter de control de un CIF: (dígito, letra) admitidos."""
    pares = sum(int(cuerpo7[i]) for i in (1, 3, 5))
    impares = sum(sum(divmod(2 * int(cuerpo7[i]), 10)) for i in (0, 2, 4, 6))
    d = (10 - (pares + impares) % 10) % 10
    return str(d), "JABCDEFGHI"[d]


def nif_valido(nif):
    if re.fullmatch(r"[0-9]{8}[A-Z]", nif):
        return LETRAS_DNI[int(nif[:8]) % 23] == nif[8]
    if re.fullmatch(r"[XYZ][0-9]{7}[A-Z]", nif):
        return LETRAS_DNI[int(str("XYZ".index(nif[0])) + nif[1:8]) % 23] == nif[8]
    if re.fullmatch(r"[KLM][0-9]{7}[A-Z]", nif) and LETRAS_DNI[int(nif[1:8]) % 23] == nif[8]:
        return True          # K/L/M: algunas fuentes usan la letra del DNI
    if re.fullmatch(r"[ABCDEFGHJKLMNPQRSUVW][0-9]{7}[0-9A-J]", nif):
        return nif[8] in control_cif(nif[1:8])
    return False


def normalizar_nif(s):
    s = re.sub(r"[ \-./_]", "", ascii_a3(s).upper())
    if s.startswith("ES") and nif_valido(s[2:]):
        s = s[2:]
    return s


# --------------------------------------------------------------------------------
#  Series y rango de documentos
# --------------------------------------------------------------------------------
def serie_y_numero(doc):
    m = re.fullmatch(r"(.*?)(\d*)", texto(doc))
    return norm(m.group(1)), (int(m.group(2)) if m.group(2) else None)


def series_candidatas(doc):
    """Valores de serie que una persona puede escribir para el documento: la serie
    completa y cada comienzo suyo que acaba justo antes de un espacio, o antes de un
    dígito que sigue a una letra ("F-2026-001" -> {"F", "F 2026"};
    "FV2026-000017" -> {"FV", "FV2026"})."""
    s, _ = serie_y_numero(doc)
    cand = {s} if s else set()
    for i in range(1, len(s)):
        if s[i] == " " or (s[i].isdigit() and s[i - 1].isalpha()):
            cand.add(s[:i])
    return cand


def serie_en_lista(doc, lista):
    return bool(series_candidatas(doc) & {norm(x) for x in lista.split(";") if norm(x)})


# --------------------------------------------------------------------------------
#  Motor del oráculo (perfil GENERAL)
# --------------------------------------------------------------------------------
ABONO = {"ABONO", "A", "NOTA DE CREDITO", "NC"}
RECTIFICATIVA = {"RECTIFICATIVA", "R"}
TICKET = {"TICKET", "T", "SIMPLIFICADA", "S", "FACTURA SIMPLIFICADA"}
NO_ANULADA = {"NO", "N", "0", "FALSO", "FALSE"}
NOMBRE_VACIO = "Clientes varios"
ABONO_SI_NEGATIVO = True
SUBTIPOS_VENTAS = {"01", "02", "03", "04", "05", "06", "08", "09"}
RETENCIONES_HABITUALES = {Decimal(x) for x in (1, 2, 7, 15, 19, 24)}

COLUMNAS = {
    "fecha": ["Fecha", "Fecha factura"], "doc": ["Nº Factura", "Numero factura", "Factura", "Documento"],
    "tipo": ["Tipo", "Tipo documento"], "cliente": ["Cliente", "Nombre cliente"], "nif": ["NIF", "CIF", "NIF cliente"],
    "cp": ["CP", "Código postal"], "base": ["Base imponible", "Base"], "piva": ["% IVA", "Tipo IVA", "IVA %"],
    "cuota": ["Cuota IVA", "Cuota"], "total": ["Total", "Total factura"], "anulada": ["Anulada"],
    "cta_cli": ["Cuenta cliente"], "desc_cli": ["Nombre cuenta cliente"], "cta_ven": ["Cuenta ventas"],
    "nom_ven": ["Nombre cuenta ventas"], "cta_iva": ["Cuenta IVA"], "pre": ["% Recargo"], "cre": ["Cuota recargo"],
    "pret": ["% Retención"], "cret": ["Cuota retención"], "fecha_op": ["Fecha operación"], "subtipo": ["Subtipo"],
    "impreso": ["Impreso"]}


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


def _col(cab, nombres):
    for nombre in nombres:
        for i, c in enumerate(cab):
            if norm(texto(c)) == norm(nombre):
                return i
    return None


def con_importe(x):
    return x["base"] != 0 or x["cuota"] != 0 or x["cre"] != 0 or x["cret"] != 0


def leer_lineas(filas, emp):
    cab = filas[0]
    C = {k: _col(cab, n) for k, n in COLUMNAS.items()}

    def g(fila, k):
        i = C[k]
        return None if i is None or i >= len(fila) else fila[i]

    L = []
    for n, fila in enumerate(filas[1:], start=2):
        if all(vacio(x) for x in fila):
            continue
        e, av = [], []                     # errores de lectura y avisos de la línea
        d = {"fila": n}
        if any(isinstance(g(fila, k), ErrorCelda) for k in C):
            e.append("celda_error")
        # fecha y documento
        d["fecha"] = parse_fecha(g(fila, "fecha"))
        if d["fecha"] is None:
            e.append("fecha")
            d["clave_fecha"] = "?" + texto(g(fila, "fecha")).upper()
        else:
            d["clave_fecha"] = f8(d["fecha"])
        d["doc"] = texto(g(fila, "doc"))
        if not d["doc"]:
            e.append("doc")
        d["fecha_op"] = d["fecha"]
        if not vacio(g(fila, "fecha_op")):
            d["fecha_op"] = parse_fecha(g(fila, "fecha_op"))
            if d["fecha_op"] is None:
                e.append("fecha_op")
        # tipo y anulación
        d["tipo_txt"] = texto(g(fila, "tipo"))
        t = norm(d["tipo_txt"])
        d["tipo"] = ("abono" if t in ABONO else "rectificativa" if t in RECTIFICATIVA
                     else "ticket" if t in TICKET else "factura")
        a = texto(g(fila, "anulada"))
        d["anulada"] = bool(a) and norm(a) not in NO_ANULADA
        # cliente
        d["cliente"], d["nif"] = texto(g(fila, "cliente")), texto(g(fila, "nif"))
        v = g(fila, "cp")
        d["cp"] = texto(v)
        if es_numero(v) and re.fullmatch(r"[0-9]{4}", d["cp"]):
            d["cp"] = "0" + d["cp"]
        # importes y porcentajes
        for k, lectura in [("base", parse_importe), ("piva", lambda x: parse_pct(x, 1)), ("cuota", parse_importe),
                           ("pre", lambda x: parse_pct(x, Decimal("0.1"))), ("cre", parse_importe),
                           ("pret", lambda x: parse_pct(x, 1)), ("cret", parse_importe)]:
            d[k], ok = lectura(g(fila, k))
            if not ok:
                e.append(k)
        if d["cret"] != 0 and d["base"] != 0 and (d["cret"] < 0) != (d["base"] < 0):
            d["cret"] = -d["cret"]
            av.append("RET_SIGNO")
        if d["cret"] != 0 and d["pret"] == 0 and d["base"] != 0:
            calc = q2(d["cret"] / d["base"] * 100)
            if calc in RETENCIONES_HABITUALES:
                d["pret"] = calc
                av.append("RET_SIN_PCT")
            else:
                e.append("ret_sin_pct")
        if C["total"] is not None and not vacio(g(fila, "total")):
            d["total"], ok = parse_importe(g(fila, "total"))
            if not ok:
                e.append("total")
        else:
            d["total"] = d["base"] + d["cuota"] + d["cre"] - d["cret"]
        # cuentas
        cc = texto(g(fila, "cta_cli"))
        if cc:
            d["cta_cli"] = cc
            d["desc_cli"] = texto(g(fila, "desc_cli")) or d["cliente"] or "Cliente " + cc
        else:
            d["cta_cli"], d["desc_cli"] = emp.cta_clientes, emp.desc_clientes
        cv = texto(g(fila, "cta_ven"))
        if cv:
            d["cta_ven"] = cv
            d["desc_ven"] = texto(g(fila, "nom_ven")) or (emp.desc_ventas if cv == emp.cta_ventas else "Ventas")
        else:
            d["cta_ven"], d["desc_ven"] = emp.cta_ventas, emp.desc_ventas
        cfg = emp.iva.get(d["piva"])
        d["cta_iva"] = texto(g(fila, "cta_iva")) or (cfg[0] if cfg else "")
        d["cta_re"] = cfg[2] if cfg else ""
        if d["cre"] != 0 and d["pre"] == 0:
            d["pre"] = cfg[1] if cfg else CERO
            if d["pre"] == 0:
                e.append("re_sin_pct")
        d["cta_ret"] = emp.cta_ret
        # subtipo e impreso
        st = texto(g(fila, "subtipo"))
        if st == "":
            d["subtipo"] = "01"
            if d["piva"] == 0 and d["base"] != 0:
                av.append("IVA0_SIN_SUBTIPO")
        elif re.fullmatch(r"[0-9]{1,2}", st) and st.zfill(2) in SUBTIPOS_VENTAS:
            d["subtipo"] = st.zfill(2)
        else:
            d["subtipo"] = "??"
            e.append("subtipo")
        im = texto(g(fila, "impreso"))
        if im == "":
            d["impreso"] = "02" if d["subtipo"] in ("03", "04") else "01"
        elif re.fullmatch(r"[0-9]{1,2}", im) and im.zfill(2) != "00":
            d["impreso"] = im.zfill(2)
        else:
            d["impreso"] = "??"
            e.append("impreso")
        d["errores"], d["avisos"] = e, av
        L.append(d)
    return L


def dentro_filtro(d, fil):
    if d["fecha"] is not None:
        if fil.get("desde") and d["fecha"] < fil["desde"]:
            return False
        if fil.get("hasta") and d["fecha"] > fil["hasta"]:
            return False
    rd, rh = fil.get("ref_desde", ""), fil.get("ref_hasta", "")
    if rd or rh:
        ref = serie_y_numero(rd or rh)[0]
        s, num = serie_y_numero(d["doc"])
        if s != ref or num is None:
            return False
        if rd and serie_y_numero(rd)[1] is not None and num < serie_y_numero(rd)[1]:
            return False
        if rh and serie_y_numero(rh)[1] is not None and num > serie_y_numero(rh)[1]:
            return False
    if fil.get("incl") and not serie_en_lista(d["doc"], fil["incl"]):
        return False
    if fil.get("excl") and serie_en_lista(d["doc"], fil["excl"]):
        return False
    return True


def tipo_documento(ls):
    t = ls[0]["tipo"]
    if (t == "factura" and ABONO_SI_NEGATIVO and all(x["tipo_txt"] == "" for x in ls)
            and sum(x["total"] for x in ls if con_importe(x)) < 0):
        return "abono"
    return t


def procesar(filas, emp, fil):
    """fil: dict con desde, hasta (date), ref_desde, ref_hasta, incl, excl (str),
    facturas, tickets, abonos (bool). Devuelve dict con lineas, exportados, excluidos
    [(doc, categoría, motivos)], avisos [(doc, categoría)] y resumen."""
    L = leer_lineas(filas, emp)
    filtradas = 0
    grupos, idx = [], {}
    for d in L:
        if not dentro_filtro(d, fil):
            filtradas += 1
            continue
        k = (d["doc"].upper(), d["clave_fecha"], d["anulada"])
        if k not in idx:
            idx[k] = len(grupos)
            grupos.append((k, []))
        grupos[idx[k]][1].append(d)
    anulados = {k[:2] for k in idx if k[2]}

    R = {"lineas": [], "excluidos": [], "avisos": [], "exportados": [], "numeros_a3": set(),
         "res": dict(leidas=0, exp=0, excl=0, filtradas=0, fact=0, tick=0, abon=0,
                     base=CERO, cuota=CERO, cuota_re=CERO, cuota_ret=CERO, total=CERO, origen=CERO)}
    admite = {"factura": fil.get("facturas", True), "ticket": fil.get("tickets", True),
              "abono": fil.get("abonos", True), "rectificativa": fil.get("abonos", True)}
    for k, ls in grupos:
        tipo = tipo_documento(ls)
        if not admite[tipo]:
            filtradas += len(ls)
            continue
        R["res"]["leidas"] += 1
        R["res"]["origen"] += sum(x["total"] for x in ls)
        _documento(ls, emp, tipo, not k[2] and k[:2] in anulados, R)
    R["res"]["filtradas"] = filtradas
    R["res"]["lineas"] = len(R["lineas"])
    return R


def _documento(ls, emp, tipo, con_anulado, R):
    doc, fecha = ls[0]["doc"], ls[0]["fecha"]

    def excluir(cat, motivos=()):
        R["excluidos"].append((doc, cat, tuple(motivos)))
        R["res"]["excl"] += 1

    def aviso(cat):
        R["avisos"].append((doc, cat))

    if any(x["anulada"] for x in ls):
        return excluir("ANULADO")
    errores = [e for x in ls for e in x["errores"]]
    if errores:
        return excluir("DATOS", errores)
    if sum(abs(x["total"]) for x in ls) == 0:
        return excluir("CERO")
    vivas = [x for x in ls if con_importe(x)]
    cero_con_total = sum(x["total"] for x in ls if not con_importe(x))
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
    clave_a3 = (documento_a3(doc).upper(), fecha)
    if clave_a3 in R["numeros_a3"]:
        return excluir("REPETIDO")
    R["numeros_a3"].add(clave_a3)

    # tipo de registro y signo de los importes en el .DAT
    abono_o_rect = tipo in ("abono", "rectificativa")
    if tipo == "rectificativa":
        factor = -1
        if total > 0:
            aviso("RECT_AUMENTA")
    elif tipo == "abono":
        factor = -1 if total < 0 else 1
        if total > 0:
            aviso("ABONO_POSITIVO")
    else:
        factor = 1
        if total < 0:
            aviso("NEGATIVA")
    if cero_con_total != 0:
        aviso("CERO_CON_TOTAL")
    if len(ascii_a3(doc)) > 10:
        aviso("DOC_LARGO")
    if con_anulado:
        aviso("CON_ANULADO")
    for x in vivas:
        for a in x["avisos"]:
            aviso(a)
    tolerancia = lambda b: Decimal("0.02") + abs(b) * Decimal("0.0005")
    for x in vivas:
        for p, cuota, cat in (("piva", "cuota", "CUOTA"), ("pre", "cre", "CUOTA_RE"), ("pret", "cret", "CUOTA_RET")):
            if x[p] > 0 and abs(q2(x["base"] * x[p] / 100) - x[cuota]) > tolerancia(x["base"]):
                aviso(cat)

    p = ls[0]
    cliente = ascii_a3(p["cliente"]) or NOMBRE_VACIO
    nif = normalizar_nif(p["nif"])
    cp = ""
    if nif:
        if not nif_valido(nif):
            aviso("NIF")
        cp = ascii_a3(p["cp"])
        if cp and not re.fullmatch(r"[0-9]{5}", cp):
            aviso("CP")
            cp = ""
    desc = doc + " " + cliente
    lineas = [reg_cabecera(emp.codigo, fecha, abono_o_rect, cuentas[0]["cli"], p["desc_cli"], doc, desc,
                           factor * total, nif, cliente if nif else "", cp, p["fecha_op"], fecha)]
    # detalle: una línea por cuenta de ventas + tipo de IVA (+ recargo, retención, subtipo, impreso)
    det = {}
    for x, c in zip(vivas, cuentas):
        k = (c["ven"], x["desc_ven"], x["piva"], c["iva"], x["pre"], c["re"], x["pret"], c["ret"], x["subtipo"],
             x["impreso"])
        acc = det.setdefault(k, [CERO, CERO, CERO, CERO])
        for j, campo in enumerate(("base", "cuota", "cre", "cret")):
            acc[j] += x[campo]
    claves = sorted(det, key=lambda k: (k[0], k[1], k[2], k[4], k[6]))
    signo = -1 if abono_o_rect else 1        # resumen: los abonos y rectificativas restan
    for n, k in enumerate(claves):
        ven, dven, piva, civa, pre, cre_cta, pret, cret_cta, st, impreso = k
        b, cu, re_, rt = (factor * v for v in det[k])
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
    R["res"]["abon" if abono_o_rect else ("tick" if p["tipo"] == "ticket" else "fact")] += 1
