"""Casos límite del perfil GENERAL de facturas emitidas: el motor VBA se ejecuta en
LibreOffice (una sola sesión para todos los escenarios) y cada documento se compara
con el oráculo independiente oraculo_facturas.py (registros 1/2 y 9, exclusiones,
avisos y resumen).

Exclusiones y avisos se comparan por documento y categoría contando repeticiones (los avisos
de línea salen una vez por línea); en las exclusiones por datos no válidos se comprueba además
que el texto del VBA menciona cada motivo que ve el oráculo.

Los fallos que corresponden a errores reales del VBA ya documentados se marcan con el
identificador del hallazgo (casos-limite-NN, lista CONOCIDOS). La prueba termina con código 1
mientras alguno siga abierto; cuando se corrija el VBA debe pasar sin tocar nada aquí."""
import datetime
import pathlib
import sys
from collections import Counter, defaultdict
from decimal import Decimal

from lo_merge import merge
import lo_runner
import oraculo_facturas as O

T = pathlib.Path(__file__).resolve().parent
DATOS = T / "datos"
SALIDA = T / "salida"
SALIDA.mkdir(exist_ok=True)
D = datetime.date

# ------------------------------------------------------------------------------
#  Empresa y tabla de IVA (idénticas para el VBA y para el oráculo)
# ------------------------------------------------------------------------------
EMP = O.Empresa(codigo="01700", digitos=8, cta_clientes="43000000", desc_clientes="Clientes varios",
                cta_ventas="70000000", desc_ventas="Ventas de mercaderías", cta_ret="47300000",
                iva={Decimal("0.00"): ("47700000", Decimal("0.00"), ""),
                     Decimal("4.00"): ("47700004", Decimal("0.50"), "47710004"),
                     Decimal("10.00"): ("47700010", Decimal("1.40"), "47710010"),
                     Decimal("21.00"): ("47700021", Decimal("5.20"), "47710021")})

# ------------------------------------------------------------------------------
#  Origen "hoja de Excel": valores con tipo (fecha, número, % con formato %)
# ------------------------------------------------------------------------------
CAB = ["Fecha", "Nº Factura", "Tipo", "Cliente", "NIF", "CP", "Base imponible", "% IVA", "Cuota IVA", "Total",
       "Cuenta cliente", "Cuenta ventas", "Nombre cuenta ventas", "% Recargo", "Cuota recargo", "% Retención",
       "Cuota retención", "Subtipo", "Anulada"]


def fila(fecha, doc, tipo, cliente, nif, cp, base, piva, cuota, total, pre=None, cre=None, pret=None, cret=None,
         st=None, anu=None):
    return [fecha, doc, tipo, cliente, nif, cp, base, piva, cuota, total, None, None, None, pre, cre, pret, cret,
            st, anu]


HOJA = [CAB,
        # recargo 5,2 % escrito en Excel como "5,2%" -> celda numérica 0,052
        fila(D(2026, 2, 1), "H-1", "Factura", "Mayorista Sur S.L.", "B41000019", "41001", 100.0, 0.21, 21.0, 126.2,
             pre=0.052, cre=5.2),
        # retención 15 % escrita como "15%" -> 0,15
        fila(D(2026, 2, 2), "H-2", "Factura", "Asesoría Técnica Ruiz", "B41000027", "41002", 1000.0, 0.21, 210.0,
             1060.0, pret=0.15, cret=150.0),
        # nº de factura numérico, IVA 10 % como 0,1
        fila(D(2026, 2, 3), 1001.0, "Factura", "Cliente Numérico", "B41000035", "41003", 1234.56, 0.1, 123.46,
             1358.02),
        fila(D(2026, 2, 4), "H-4", "Abono", "Mayorista Sur S.L.", "B41000019", "41001", -50.0, 0.21, -10.5, -60.5),
        # IVA 4 % como 0,04 y recargo 0,5 % escrito como número 0,5: NO es formato % (límite 0,1)
        fila(D(2026, 2, 5), "H-5", "Factura", "Farmacia Sur", "B41000050", "41005", 100.0, 0.04, 4.0, 104.5,
             pre=0.5, cre=0.5),
        # CP en celda numérica (8001 -> 08001) y subtipo numérico 2 -> 02
        fila(D(2026, 2, 6), "H-6", "Factura", "Exportaciones Norte", "B08000010", 8001.0, 200.0, 0.0, 0.0, 200.0,
             st=2.0),
        # celda con error de fórmula (#N/D) en una columna de texto del perfil
        fila(D(2026, 2, 7), "H-7", "Factura", O.ErrorCelda(), "B41000076", "41007", 100.0, 0.21, 21.0, 121.0)]

# ------------------------------------------------------------------------------
#  Escenarios: (nombre, origen, filtros)
# ------------------------------------------------------------------------------
ESCENARIOS = [
    ("completo", "general_casos.csv", {}),
    ("f_fechas", "general_filtros.csv", dict(desde=D(2026, 1, 10), hasta=D(2026, 1, 20))),
    ("f_desde", "general_filtros.csv", dict(desde=D(2026, 1, 21))),
    ("f_rango", "general_filtros.csv", dict(ref_desde="INV102", ref_hasta="INV104")),
    ("f_rango_desde", "general_filtros.csv", dict(ref_desde="INV104")),
    ("f_series_incl", "general_filtros.csv", dict(incl="INV;CRN")),
    ("f_series_excl", "general_filtros.csv", dict(excl="REC;2026")),
    ("f_solo_facturas", "general_filtros.csv", dict(tickets=False, abonos=False)),
    ("f_tickets_abonos", "general_filtros.csv", dict(facturas=False)),
    ("f_combinado", "general_filtros.csv", dict(desde=D(2026, 1, 10), hasta=D(2026, 1, 20), incl="INV",
                                                tickets=False, abonos=False)),
    ("sintipo_todo", "general_sintipo.csv", {}),
    ("sintipo_sin_abonos", "general_sintipo.csv", dict(abonos=False)),
    ("sintipo_solo_abonos", "general_sintipo.csv", dict(facturas=False, tickets=False)),
    ("total_vacio", "general_total_vacio.csv", {}),
    ("plantilla_excl_R", "general_casos.csv", dict(excl="R")),
    ("plantilla_incl_F", "general_casos.csv", dict(incl="F")),
    # "FV" vale para FV2026-*, "T" para T-0001 y "F-20" no vale para "F 2026" (acaba en dígito)
    ("plantilla_incl_FV", "general_casos.csv", dict(incl="FV;F-20;T")),
    ("reglas", "general_reglas.csv", {}),
    ("reglas_sin_abonos", "general_reglas.csv", dict(abonos=False)),
    ("reglas_solo_abonos", "general_reglas.csv", dict(facturas=False, tickets=False)),
    ("hoja_excel", "HOJA", {}),
]

# Fallos conocidos (errores del VBA documentados): hallazgo -> {escenario: documentos}.
# "*" = diferencias del resumen del escenario; "F-2026-*" = prefijo; "#colision" = nº repetido en a3.
# Cualquier otra diferencia se informa como NUEVA.
# Corregidos en el VBA (ya no se esperan): casos-limite-01 (total vacío), 02 (filtro de tipos por
# documento), 03 (% con formato % de Excel), 04 (series "F" / "R" de "F-2026-001"), 05 (impreso 02 en
# subtipo 03) y 06 (nº de factura repetido en a3).
CONOCIDOS = {}


def hallazgo_conocido(escenario, doc):
    for h, por_esc in CONOCIDOS.items():
        for patron in por_esc.get(escenario, []):
            if patron == doc or (patron.endswith("-*") and doc.startswith(patron[:-1])):
                return h
    return None


# ------------------------------------------------------------------------------
#  Generación del código VBA de la prueba
# ------------------------------------------------------------------------------
def vb(v):
    if isinstance(v, O.ErrorCelda):
        return f"CVErr({v.codigo})"
    if isinstance(v, str):
        return '"' + v.replace('"', '""') + '"'
    if isinstance(v, datetime.date):
        return f"DateSerial({v.year}, {v.month}, {v.day})"
    if isinstance(v, bool):
        return "True" if v else "False"
    return repr(v)


def codigo_vba():
    L = []
    L.append("Private Sub FiltrosReiniciar(ByRef f As TFiltros)")
    L.append("    f.UsarDesde = False: f.UsarHasta = False: f.RefDesde = \"\": f.RefHasta = \"\"")
    L.append("    f.SeriesIncluir = \"\": f.SeriesExcluir = \"\": f.ExcluirDescuadrados = False")
    L.append("    f.InclFacturas = True: f.InclTickets = True: f.InclAbonos = True")
    L.append("End Sub")
    L.append("Private Function MatrizHoja() As Variant")
    L.append("    Dim m() As Variant")
    L.append(f"    ReDim m(1 To {len(HOJA)}, 1 To {len(CAB)})")
    for r, f in enumerate(HOJA, start=1):
        for c, v in enumerate(f, start=1):
            if v is not None:
                L.append(f"    m({r}, {c}) = {vb(v)}")
    L.append("    MatrizHoja = m")
    L.append("End Function")
    L.append("Private Function Escenario(ByVal origen As String, ByVal salida As String, ByRef fil As TFiltros) As String")
    L.append("    On Error GoTo EH")
    L.append("    Dim emp As TEmpresa, per As TPerfil, res As TResumen, msg As String, datos As Variant, s As String, i As Long")
    L.append(f"    emp.Codigo = {vb(EMP.codigo)}: emp.Digitos = {EMP.digitos}")
    L.append(f"    emp.CtaClientes = {vb(EMP.cta_clientes)}: emp.DescClientes = {vb(EMP.desc_clientes)}")
    L.append(f"    emp.CtaVentas = {vb(EMP.cta_ventas)}: emp.DescVentas = {vb(EMP.desc_ventas)}")
    L.append(f"    emp.CtaRetencion = {vb(EMP.cta_ret)}")
    L.append("    If Not PerfilPredefinido(\"GENERAL\", per) Then")
    L.append("        Escenario = \"ERR perfil\"")
    L.append("        Exit Function")
    L.append("    End If")
    L.append("    IvaReiniciar")
    for p, (cta, pre, cta_re) in EMP.iva.items():
        L.append(f"    IvaAgregar {float(p)!r}, {vb(cta)}, {float(pre)!r}, {vb(cta_re)}")
    L.append("    If origen = \"HOJA\" Then")
    L.append("        datos = MatrizHoja()")
    L.append("    Else")
    L.append("        datos = ParsearCSV(LeerFicheroTexto(origen), per.Separador)")
    L.append("    End If")
    L.append("    If Not FacturasProcesar(emp, per, fil, datos, res, msg) Then")
    L.append("        Escenario = \"ERR \" & msg")
    L.append("        Exit Function")
    L.append("    End If")
    L.append("    If gNDat > 0 Then EscribirDat salida")
    L.append("    s = \"RES\" & vbTab & res.UnidadesLeidas & vbTab & res.UnidadesExportadas & vbTab & res.UnidadesExcluidas & _")
    L.append("        vbTab & res.FilasFiltradas & vbTab & res.NumFacturas & vbTab & res.NumTickets & vbTab & res.NumAbonos & _")
    L.append("        vbTab & gNDat & vbTab & ImporteA3(res.BaseImp) & vbTab & ImporteA3(res.Cuota) & vbTab & ImporteA3(res.CuotaRE) & _")
    L.append("        vbTab & ImporteA3(res.CuotaRet) & vbTab & ImporteA3(res.Total) & vbTab & ImporteA3(res.TotalOrigen)")
    L.append("    For i = 1 To gNInc")
    L.append("        s = s & Chr(10) & \"INC\" & vbTab & gInc(i).Gravedad & vbTab & gInc(i).Referencia & vbTab & gInc(i).Fila & vbTab & gInc(i).Texto")
    L.append("    Next i")
    L.append("    Escenario = s")
    L.append("    Exit Function")
    L.append("EH:")
    L.append("    Escenario = \"ERR \" & Err.Number & \" \" & Err.Description")
    L.append("End Function")
    L.append("Function TestGeneral() As String")
    L.append("    Dim fil As TFiltros, s As String")
    for nombre, origen, f in ESCENARIOS:
        L.append("    FiltrosReiniciar fil")
        if f.get("desde"):
            L.append(f"    fil.UsarDesde = True: fil.FechaDesde = {vb(f['desde'])}")
        if f.get("hasta"):
            L.append(f"    fil.UsarHasta = True: fil.FechaHasta = {vb(f['hasta'])}")
        for k, campo in [("ref_desde", "RefDesde"), ("ref_hasta", "RefHasta"), ("incl", "SeriesIncluir"),
                         ("excl", "SeriesExcluir")]:
            if f.get(k):
                L.append(f"    fil.{campo} = {vb(f[k])}")
        for k, campo in [("facturas", "InclFacturas"), ("tickets", "InclTickets"), ("abonos", "InclAbonos")]:
            if k in f:
                L.append(f"    fil.{campo} = {vb(f[k])}")
        ruta = "HOJA" if origen == "HOJA" else str(DATOS / origen)
        L.append(f"    s = s & \"##ESC\" & vbTab & {vb(nombre)} & Chr(10) & Escenario({vb(ruta)}, "
                 f"{vb(str(dat_de(nombre)))}, fil) & Chr(10)")
    L.append("    TestGeneral = s")
    L.append("End Function")
    return "\n".join(L)


def dat_de(nombre):
    return SALIDA / f"SUENLACE_general_{nombre}.DAT"


# ------------------------------------------------------------------------------
#  Comparación
# ------------------------------------------------------------------------------
CAMPOS_CAB = [(1, 1, "formato"), (2, 6, "empresa"), (7, 14, "fecha"), (15, 15, "tipo registro"),
              (16, 27, "cuenta cliente"), (28, 57, "desc. cuenta"), (58, 58, "tipo factura"), (59, 68, "documento"),
              (69, 69, "linea"), (70, 99, "descripcion"), (100, 113, "total"), (114, 175, "reserva"),
              (176, 189, "NIF"), (190, 229, "nombre"), (230, 234, "CP"), (235, 236, "reserva"),
              (237, 244, "fecha operacion"), (245, 252, "fecha factura"), (253, 254, "moneda/generado")]
CAMPOS_DET = [(1, 1, "formato"), (2, 6, "empresa"), (7, 14, "fecha"), (15, 15, "tipo registro"),
              (16, 27, "cuenta ventas"), (28, 57, "desc. cuenta"), (58, 58, "cargo/abono"), (59, 68, "documento"),
              (69, 69, "linea"), (70, 99, "descripcion"), (100, 101, "subtipo"), (102, 115, "base"),
              (116, 120, "% IVA"), (121, 134, "cuota"), (135, 139, "% recargo"), (140, 153, "cuota recargo"),
              (154, 158, "% retencion"), (159, 172, "cuota retencion"), (173, 174, "impreso"),
              (175, 175, "sujeta"), (176, 176, "415"), (177, 191, "marcas/reserva"), (192, 203, "cuenta IVA"),
              (204, 215, "cuenta recargo"), (216, 227, "cuenta retencion"), (228, 252, "reserva"),
              (253, 254, "moneda/generado")]


def diff_registro(a, b):
    campos = CAMPOS_DET if b[14] == "9" else CAMPOS_CAB
    return [f"{n} [{a[i - 1:j]!r} en vez de {b[i - 1:j]!r}]" for i, j, n in campos if a[i - 1:j] != b[i - 1:j]]


def por_documento(lineas):
    """Agrupa registros por cabecera: clave (documento, fecha, nº de aparición)."""
    out, cont = {}, Counter()
    actual = None
    for l in lineas:
        if l[14] in "12":
            k0 = (l[58:68].rstrip(), l[6:14])
            cont[k0] += 1
            actual = k0 + (cont[k0],)
            out[actual] = []
        out[actual].append(l)
    return out


CAT_EXCL = [("Documento anulado", "ANULADO"), ("Datos no válidos", "DATOS"), ("Importe 0,00", "CERO"),
            ("Sin líneas", "SINLINEAS"), ("Tipo de IVA no configurado", "IVA"), ("Cuenta no válida", "CUENTA"),
            ("Descuadre", "DESCUADRE"), ("repetido en a3", "REPETIDO")]
# el orden importa: gana la primera clave que aparece en el texto
CAT_AVISO = [("Rectificativa que aumenta la factura original", "RECT_AUMENTA"),
             ("Abono con importes en positivo", "ABONO_POSITIVO"), ("Factura con total negativo", "NEGATIVA"),
             ("base y cuota a 0 pero con total", "CERO_CON_TOTAL"), ("más de 10 caracteres", "DOC_LARGO"),
             ("Mismo nº y fecha que un documento anulado", "CON_ANULADO"),
             ("La cuota de recargo", "CUOTA_RE"), ("La retención", "CUOTA_RET"), ("no cuadra con base", "CUOTA"),
             ("retención con el signo cambiado", "RET_SIGNO"), ("retención sin %", "RET_SIN_PCT"),
             ("IVA 0% sin subtipo", "IVA0_SIN_SUBTIPO"), ("no es un NIF español válido", "NIF"),
             ("Código postal no español", "CP")]
# motivo de exclusión por datos del oráculo -> texto que debe aparecer en la incidencia del VBA
MOTIVOS_DATOS = {"celda_error": "error de fórmula", "fecha": "fecha no válida", "doc": "número de factura vacío",
                 "fecha_op": "fecha de operación no válida", "base": "base no numérica",
                 "piva": "% IVA no numérico", "cuota": "cuota no numérica", "pre": "% recargo no numérico",
                 "cre": "cuota de recargo no numérica", "pret": "% retención no numérico",
                 "cret": "retención no numérica", "total": "total no numérico",
                 "ret_sin_pct": "no es un tipo habitual", "re_sin_pct": "recargo de equivalencia sin %",
                 "subtipo": "subtipo no válido", "impreso": "impreso no válido"}


def categoria(texto, tabla):
    for clave, cat in tabla:
        if clave in texto:
            return cat
    return "OTRO: " + texto


def categorias_aviso(texto):
    """Un aviso de lectura puede juntar varios separados por "; " (uno por línea del origen)."""
    return [categoria(t, CAT_AVISO) for t in texto.split("; ")]


def comparar(nombre, origen, f, bloque):
    """Devuelve lista de (documento o '*', mensaje)."""
    fallos = []
    if bloque[0].startswith("ERR"):
        return [("*", "el VBA devuelve " + bloque[0])]
    res_vba = bloque[0].split("\t")
    incs = [l.split("\t") for l in bloque[1:] if l.startswith("INC")]
    filas = HOJA if origen == "HOJA" else O.leer_csv(DATOS / origen)
    R = O.procesar(filas, EMP, f)
    # registros
    ruta = dat_de(nombre)
    obt_bytes = ruta.read_bytes() if ruta.exists() else b""
    if obt_bytes and not obt_bytes.endswith(b"\r\n"):
        fallos.append(("*", "el .DAT no termina en CR/LF"))
    obt = obt_bytes.decode("ascii").split("\r\n")[:-1] if obt_bytes else []
    for l in obt:
        if len(l) != 254:
            fallos.append(("*", f"registro de {len(l)} posiciones"))
    esp = R["lineas"]
    a, b = por_documento(obt), por_documento(esp)
    nombres = {(O.documento_a3(doc), O.f8(fe)): doc for doc, fe, _ in R["exportados"]}
    for k in list(dict.fromkeys(list(b) + list(a))):
        doc = nombres.get(k[:2], k[0])
        if k not in a:
            fallos.append((doc, "FALTA en el .DAT del VBA (el oráculo lo exporta)"))
        elif k not in b:
            fallos.append((doc, "SOBRA en el .DAT del VBA: " + " / ".join(x[14] for x in a[k]) + " registros"))
        elif a[k] != b[k]:
            if len(a[k]) != len(b[k]):
                fallos.append((doc, f"{len(a[k])} registros en vez de {len(b[k])}"))
            for n, (x, y) in enumerate(zip(a[k], b[k])):
                if x != y:
                    fallos.append((doc, f"registro {n + 1} (tipo {y[14]}): " + "; ".join(diff_registro(x, y))))
    if obt != esp and not any(d != "*" for d, _ in fallos):
        fallos.append(("*", "mismo contenido pero distinto orden de registros"))
    # dos documentos distintos con el mismo nº y fecha en a3
    for k, n in Counter((l[58:68], l[6:14]) for l in obt if l[14] in "12").items():
        if n > 1:
            fallos.append(("#colision", f"{n} facturas distintas llegan a a3 con el mismo nº {k[0].strip()!r} y fecha {k[1]}"))
    # exclusiones (documento + categoría, con repeticiones) y motivo de las de datos no válidos
    excl_vba = [(i[2], categoria(i[4], CAT_EXCL), i[4]) for i in incs if i[1] == "EXCLUIDO"]
    obt_c = Counter((d, c) for d, c, _ in excl_vba)
    esp_c = Counter((d, c) for d, c, _ in R["excluidos"])
    for doc, cat in sorted((obt_c - esp_c) + (esp_c - obt_c)):
        fallos.append((doc, f"exclusión {cat} " + ("de más" if obt_c[(doc, cat)] > esp_c[(doc, cat)] else "que falta")))
    for doc, cat, motivos in R["excluidos"]:
        textos = [t for d, c, t in excl_vba if (d, c) == (doc, cat)]
        for m in motivos:
            if textos and not any(MOTIVOS_DATOS[m] in t for t in textos):
                fallos.append((doc, f"la exclusión no menciona '{MOTIVOS_DATOS[m]}': {textos[0]}"))
    # avisos (documento + categoría, con repeticiones: uno por línea en los avisos de línea)
    obt_c = Counter((i[2], c) for i in incs if i[1] == "AVISO" for c in categorias_aviso(i[4]))
    esp_c = Counter(R["avisos"])
    for doc, cat in sorted((obt_c - esp_c) + (esp_c - obt_c)):
        n_obt, n_esp = obt_c[(doc, cat)], esp_c[(doc, cat)]
        fallos.append((doc, f"aviso {cat} " + (f"de más ({n_obt} en vez de {n_esp})" if n_obt > n_esp
                                               else f"que falta ({n_obt} en vez de {n_esp})")))
    # resumen
    r = R["res"]
    nombres_res = ["leidas", "exp", "excl", "filtradas", "fact", "tick", "abon", "lineas",
                   "base", "cuota", "cuota_re", "cuota_ret", "total", "origen"]
    for n, v in zip(nombres_res, res_vba[1:]):
        e = r[n]
        e = O.imp(e) if isinstance(e, Decimal) else str(e)
        if v != e:
            fallos.append(("*", f"resumen {n} = {v} en vez de {e}"))
    return fallos


def main():
    for nombre, _, _ in ESCENARIOS:
        if dat_de(nombre).exists():
            dat_de(nombre).unlink()
    code = merge(["modA3Nucleo.bas", "modA3Config.bas", "modA3Facturas.bas"], codigo_vba())
    (SALIDA / "lo_merged_general.bas").write_text(code, encoding="utf-8")
    r = lo_runner.run([("Merged", code)], "Merged.TestGeneral")
    if r is None:
        sys.exit("La macro no se ha ejecutado (error de compilación en LibreOffice)")
    bloques = defaultdict(list)
    actual = None
    for l in r.split("\n"):
        if l.startswith("##ESC\t"):
            actual = l.split("\t")[1]
        elif actual and l:
            bloques[actual].append(l)
    abiertos, nuevos = set(), 0
    for nombre, origen, f in ESCENARIOS:
        fallos = comparar(nombre, origen, f, bloques.get(nombre, ["ERR sin salida"]))
        if not fallos:
            print(f"OK     {nombre}")
            continue
        print(f"FALLO  {nombre}")
        for doc, msg in fallos:
            h = hallazgo_conocido(nombre, doc)
            if h:
                abiertos.add(h)
            else:
                nuevos += 1
            print(f"         {doc}: {msg}" + (f"   <- {h}" if h else "   <- NUEVO"))
    print()
    if abiertos:
        print("Hallazgos abiertos (errores del VBA documentados): " + ", ".join(sorted(abiertos)))
    if nuevos:
        print(f"{nuevos} diferencia(s) SIN documentar")
    print("GENERAL OK: todo coincide con el oráculo" if not (abiertos or nuevos) else "GENERAL CON DIFERENCIAS")
    sys.exit(0 if not (abiertos or nuevos) else 1)


if __name__ == "__main__":
    main()
