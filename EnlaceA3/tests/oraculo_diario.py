"""Oráculo independiente (Python) del registro tipo 0, para comparar con el VBA."""
import unicodedata
def txt(s, n):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = s.replace("–", "-")
    return " ".join(s.split())[:n].ljust(n)
def imp(x):
    return ("-" if x < 0 else "+") + f"{abs(x):013.2f}"
def reg0(emp, fecha, cta, desc_cta, dh, doc, imu, desc, importe):
    r = "4" + emp + fecha + "0" + txt(cta, 12) + txt(desc_cta, 30) + dh + txt(doc, 10) + imu + txt(desc, 30) + imp(importe) + " " * 138 + " EN"
    assert len(r) == 254, len(r)
    return r
EMP = "01696"
def esperado():
    L = []
    def asiento(fecha, lineas):
        n = len(lineas)
        for k, (cta, nom, desc, doc, dh, im) in enumerate(lineas):
            imu = "I" if k == 0 else ("U" if k == n - 1 else "M")
            L.append(reg0(EMP, fecha, cta, nom, dh, doc, imu, desc, im))
    asiento("20260101", [("570000000", "Caja, euros", "ASIENTO DE APERTURA", "", "D", 1500.00),
                         ("572000001", "BANCO BILBAO VIZCAYA ARGENTARIA S.A. - ES65 0182", "ASIENTO DE APERTURA", "", "D", 8500.00),
                         ("100000000", "Capital social", "ASIENTO DE APERTURA", "", "H", 10000.00)])
    asiento("20260102", [("629000000", "Otros servicios", "Pago nómina señor Muñoz - enero", "F-001", "D", 248.02),
                         ("572000001", "Banco", "Pago nómina señor Muñoz - enero", "F-001", "H", 248.02)])
    asiento("20260102", [("600000000", "Compras", "Regularización con Debe y Haber", "", "D", 100.00),
                         ("600000000", "Compras", "Regularización con Debe y Haber", "", "H", 100.00)])
    asiento("20260103", [("572000005", "Banco 5", "Cuenta con punto", "", "D", 75.50),
                         ("430000001", "Cliente A", "Cuenta con punto", "", "H", 75.50)])
    asiento("20260104", [("572000001", "Banco", "Con linea a cero en medio", "", "D", 20.00),
                         ("400000001", "Proveedor", "Con linea a cero en medio", "DOC7", "H", 20.00)])
    asiento("20260104", [("572000001", "Banco", "Ultima linea a cero", "", "H", 33.33),
                         ("410000001", "Acreedor", "Ultima linea a cero", "", "D", 33.33)])
    asiento("20260105", [("572000001", "Banco", "Mismo asiento otra fecha", "", "D", 1.00),
                         ("570000000", "Caja", "Mismo asiento otra fecha", "", "H", 1.00)])
    asiento("20260106", [("572000001", "Banco", "Mismo asiento otra fecha", "", "D", 2.00),
                         ("570000000", "Caja", "Mismo asiento otra fecha", "", "H", 2.00)])
    return L
