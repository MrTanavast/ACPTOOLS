Attribute VB_Name = "modA3Nucleo"
' =====================================================================
'  ENLACE CONTABLE A3  ·  ACP ASOCIADOS (Córdoba)
'  Módulo NÚCLEO: tipos de datos, formato de los campos del fichero
'  SUENLACE.DAT, conversión de textos / importes / fechas, lectura de
'  ficheros CSV y escritura del .DAT.
'
'  Formato del fichero (Wolters Kluwer a3 - "Enlace contable de entrada"):
'    - Formato "4": registros de 254 posiciones + CR/LF (256 bytes).
'    - Texto ASCII puro (sin tildes ni eñes).
'    - Registro 0 = apunte sin IVA (diarios)
'    - Registro 1/2 = cabecera de factura / abono  +  registro 9 = detalle IVA
'
'  Todo el cálculo de importes se hace con el tipo Currency (exacto,
'  sin errores de coma flotante) y todos los formatos se construyen a
'  mano para no depender de la configuración regional de Windows.
' =====================================================================
Option Explicit

Public Const A3_VERSION As String = "4.0"
Public Const A3_LONG_REGISTRO As Long = 254
Public Const A3_NOMBRE_FICHERO As String = "SUENLACE.DAT"

' --- Colores corporativos ACP Asociados (valor BGR de VBA) -------------
Public Const COLOR_ACP As Long = &H986114          ' RGB(20, 97, 152)  azul ACP
Public Const COLOR_ACP_OSCURO As Long = &H633F0D   ' RGB(13, 63, 99)
Public Const COLOR_ACP_MEDIO As Long = &HEAD9C5    ' RGB(197, 217, 234)
Public Const COLOR_ACP_CLARO As Long = &HF7F0E8    ' RGB(232, 240, 247)
Public Const COLOR_BLANCO As Long = &HFFFFFF
Public Const COLOR_TEXTO As Long = &H333333        ' RGB(51, 51, 51)
Public Const COLOR_GRIS As Long = &H7F7F7F         ' RGB(127, 127, 127)
Public Const COLOR_OK As Long = &H327D2E           ' RGB(46, 125, 50)
Public Const COLOR_AVISO As Long = &H26CED          ' RGB(237, 108, 2)
Public Const COLOR_ERROR As Long = &H2828C6        ' RGB(198, 40, 40)
Public Const COLOR_OK_CLARO As Long = &HE9F5E8     ' RGB(232, 245, 233)
Public Const COLOR_AVISO_CLARO As Long = &HE0F3FF  ' RGB(255, 243, 224)
Public Const COLOR_ERROR_CLARO As Long = &HEEEBFF  ' RGB(255, 235, 238)

' --- Modos de enlace ------------------------------------------------------
Public Const MODO_DIARIO As Integer = 1
Public Const MODO_EMITIDAS As Integer = 2
Public Const MODO_RECIBIDAS As Integer = 3     ' preparado, aún no disponible

' --- Origen de los datos -------------------------------------------------
Public Const ORIGEN_HOJA As Integer = 1
Public Const ORIGEN_CSV As Integer = 2

' --- Gravedad de las incidencias ------------------------------------------
Public Const INC_EXCLUIDO As String = "EXCLUIDO"
Public Const INC_AVISO As String = "AVISO"
Public Const INC_INFO As String = "INFO"

' =====================================================================
'  TIPOS DE DATOS COMPARTIDOS
'  (solo campos simples: así el código es portable y fácil de probar)
' =====================================================================

' Configuración de una empresa (hoja EMPRESAS)
Public Type TEmpresa
    Codigo As String            ' 5 dígitos, p.ej. "01692"
    Nombre As String
    Digitos As Integer          ' longitud del plan de cuentas (6 a 12)
    Perfil As String            ' perfil del libro de facturas (hoja PERFILES)
    CtaClientes As String
    DescClientes As String
    CtaVentas As String         ' cuenta de ventas general
    DescVentas As String
    CtaVentasAlt As String      ' cuenta de ventas si la columna "selectora" tiene dato
    DescVentasAlt As String
    CtaRetencion As String
    CarpetaSalida As String
    Existe As Boolean
    Fila As Long                ' fila en la hoja EMPRESAS (0 si no existe)
End Type

' Perfil de lectura de un libro de facturas (hoja PERFILES)
Public Type TPerfil
    Nombre As String
    Campos As String            ' nombres de columna de cada campo, separados por vbTab
    Separador As String         ' "auto", ";", ",", "TAB", "|"
    SepDecimal As String        ' "auto", ".", ","
    FormatoFecha As String      ' "DMA" (dd/mm/aaaa), "MDA" o "AMD"
    ValoresAbono As String      ' valores de la columna TIPO que indican abono (separados por ;)
    ValoresTicket As String     ' valores de la columna TIPO que indican ticket
    ValoresRectificativa As String ' rectificativas: tipo 2 que toma el signo del origen
    NombreVacio As String       ' nombre de cliente si viene vacío
    AbonoSiNegativo As Boolean  ' sin columna TIPO: total negativo = abono
    FilaCabecera As Long
End Type

' Filtros elegidos en el formulario
Public Type TFiltros
    UsarDesde As Boolean
    FechaDesde As Date
    UsarHasta As Boolean
    FechaHasta As Date
    RefDesde As String          ' documento / asiento inicial (vacío = sin límite)
    RefHasta As String          ' documento / asiento final
    SeriesIncluir As String     ' series separadas por ; (vacío = todas)
    SeriesExcluir As String
    InclFacturas As Boolean
    InclTickets As Boolean
    InclAbonos As Boolean
    ExcluirDescuadrados As Boolean
End Type

' Parámetros de una ejecución
Public Type TParametros
    Empresa As String
    Digitos As Integer
    Modo As Integer
    Perfil As String
    OrigenTipo As Integer
    OrigenLibro As String
    OrigenHoja As String
    OrigenRuta As String
    OrigenDescripcion As String
    CarpetaSalida As String
    GenerarControl As Boolean
    CopiaSeguridad As Boolean
End Type

' Resumen de una ejecución
Public Type TResumen
    FilasLeidas As Long
    FilasFiltradas As Long      ' filas que el filtro deja fuera
    UnidadesLeidas As Long      ' documentos o asientos dentro del filtro
    UnidadesExportadas As Long
    UnidadesExcluidas As Long
    LineasDat As Long
    NumFacturas As Long
    NumTickets As Long
    NumAbonos As Long
    BaseImp As Currency
    Cuota As Currency
    CuotaRE As Currency
    CuotaRet As Currency
    Total As Currency
    TotalOrigen As Currency     ' total de las filas de origen dentro del filtro
    TotalExcluido As Currency
    AjusteSignoAbonos As Currency
    Debe As Currency
    Haber As Currency
    HayFechas As Boolean
    FechaMin As Date
    FechaMax As Date
    Avisos As Long
    Excluidos As Long
End Type

' Incidencia (se listan en el Excel de control)
Public Type TIncidencia
    TieneFecha As Boolean
    Fecha As Date
    Referencia As String        ' documento o asiento
    Fila As Long                ' fila del origen (0 si no aplica)
    Gravedad As String          ' EXCLUIDO / AVISO / INFO
    Texto As String
    Tratamiento As String
End Type

' --- Almacenes globales de la ejecución en curso -----------------------
Public gDat() As String          ' líneas del SUENLACE.DAT (sin CR/LF)
Public gNDat As Long
Public gInc() As TIncidencia
Public gNInc As Long

Private mMapaAscii() As String
Private mMapaListo As Boolean

' =====================================================================
'  ALMACÉN DE LÍNEAS DAT E INCIDENCIAS
' =====================================================================
' Pone a cero un resumen (sin asignar tipos completos: portable a LibreOffice)
Public Sub ResumenReiniciar(ByRef r As TResumen)
    r.FilasLeidas = 0
    r.FilasFiltradas = 0
    r.UnidadesLeidas = 0
    r.UnidadesExportadas = 0
    r.UnidadesExcluidas = 0
    r.LineasDat = 0
    r.NumFacturas = 0
    r.NumTickets = 0
    r.NumAbonos = 0
    r.BaseImp = 0
    r.Cuota = 0
    r.CuotaRE = 0
    r.CuotaRet = 0
    r.Total = 0
    r.TotalOrigen = 0
    r.TotalExcluido = 0
    r.AjusteSignoAbonos = 0
    r.Debe = 0
    r.Haber = 0
    r.HayFechas = False
    r.FechaMin = 0
    r.FechaMax = 0
    r.Avisos = 0
    r.Excluidos = 0
End Sub

Public Sub DatReiniciar()
    gNDat = 0
    ReDim gDat(1 To 256)
End Sub

Public Sub DatAgregar(ByVal linea As String)
    If Len(linea) <> A3_LONG_REGISTRO Then
        Err.Raise vbObjectError + 901, "DatAgregar", _
            "Error interno: registro de " & Len(linea) & " posiciones (deben ser " & A3_LONG_REGISTRO & ")." & vbCrLf & linea
    End If
    gNDat = gNDat + 1
    If gNDat > UBound(gDat) Then
        ReDim Preserve gDat(1 To UBound(gDat) * 2)
    End If
    gDat(gNDat) = linea
End Sub

Public Sub IncReiniciar()
    gNInc = 0
    ReDim gInc(1 To 64)
End Sub

Public Sub IncAgregar(ByVal tieneFecha As Boolean, ByVal fecha As Date, ByVal referencia As String, _
                      ByVal fila As Long, ByVal gravedad As String, ByVal texto As String, _
                      ByVal tratamiento As String)
    gNInc = gNInc + 1
    If gNInc > UBound(gInc) Then
        ReDim Preserve gInc(1 To UBound(gInc) * 2)
    End If
    gInc(gNInc).TieneFecha = tieneFecha
    gInc(gNInc).Fecha = fecha
    gInc(gNInc).Referencia = referencia
    gInc(gNInc).Fila = fila
    gInc(gNInc).Gravedad = gravedad
    gInc(gNInc).Texto = texto
    gInc(gNInc).Tratamiento = tratamiento
End Sub

Public Function IncContar(ByVal gravedad As String) As Long
    Dim i As Long, n As Long
    For i = 1 To gNInc
        If gInc(i).Gravedad = gravedad Then n = n + 1
    Next i
    IncContar = n
End Function

' =====================================================================
'  TEXTOS
' =====================================================================

' Convierte un texto a ASCII imprimible: quita tildes, eñes y símbolos,
' cambia saltos de línea / tabuladores por espacios y colapsa espacios.
Public Function AsciiA3(ByVal s As String) As String
    Dim i As Long, c As Long, r As String, out As String, n As Long
    If Not mMapaListo Then CargarMapaAscii
    n = Len(s)
    out = ""
    For i = 1 To n
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        Select Case c
            Case 32 To 126
                r = ChrW(c)
            Case 0 To 31, 127 To 159
                r = " "
            Case 160 To 591
                r = mMapaAscii(c - 160)
            Case 8208 To 8213                       ' guiones tipográficos
                r = "-"
            Case 8216, 8217, 8218, 8219              ' comillas simples
                r = "'"
            Case 8220, 8221, 8222, 8223              ' comillas dobles
                r = """"
            Case 8230                                ' puntos suspensivos
                r = "..."
            Case 8364                                ' símbolo del euro
                r = "EUR"
            Case Else
                r = ""
        End Select
        out = out & r
    Next i
    AsciiA3 = ColapsarEspacios(out)
End Function

' Quita espacios al principio y al final y deja un solo espacio entre palabras.
Public Function ColapsarEspacios(ByVal s As String) As String
    s = Trim$(s)
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    ColapsarEspacios = s
End Function

' Campo de texto para el fichero: ASCII, recortado y alineado a la izquierda.
Public Function TextoA3(ByVal s As String, ByVal ancho As Long) As String
    TextoA3 = Left$(AsciiA3(s) & Space$(ancho), ancho)
End Function

' Normaliza un texto para comparar (cabeceras, series, tipos...):
' mayúsculas, sin tildes, "%" -> " PCT ", solo letras/números/espacios.
Public Function Normalizar(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase$(AsciiA3(Replace(s, "%", " PCT ")))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
        Else
            out = out & " "
        End If
    Next i
    Normalizar = ColapsarEspacios(out)
End Function

' ¿Está "valor" en la lista "a;b;c"? (comparación normalizada)
Public Function EnLista(ByVal valor As String, ByVal lista As String) As Boolean
    Dim partes() As String, i As Long, v As String
    v = Normalizar(valor)
    If v = "" Or Trim$(lista) = "" Then Exit Function
    partes = Split(lista, ";")
    For i = LBound(partes) To UBound(partes)
        If Normalizar(partes(i)) = v Then
            EnLista = True
            Exit Function
        End If
    Next i
End Function

Private Sub CargarMapaAscii()
    ' Transliteración de los caracteres Unicode 160..591 (Latin-1 y Latin
    ' Extended-A/B). Cada posición separada por "|". Posición vacía = se elimina.
    Dim s As String, partes() As String, i As Long
    s = ""
    s = s & " ||||||||||a|""|||||||2|3||||.||1|o|""|14|12|34||A|A|A|A|A|A|AE|C|E|E|E|E|I|I|I|I|D|N|O|O|O|O|O||O|U|U|U|U|Y|TH|ss|a|a|a|a|a|a|ae|c|e|e|e|e|i|i|i|i|d|n|o|o|o|o|o||o|u|u|u|u|y|th|y|"
    s = s & "A|a|A|a|A|a|C|c|C|c|C|c|C|c|D|d|D|d|E|e|E|e|E|e|E|e|E|e|G|g|G|g|G|g|G|g|H|h|||I|i|I|i|I|i|I|i|I|i|IJ|ij|J|j|K|k||L|l|L|l|L|l|L|l|L|l|N|n|N|n|N|n|n|||O|o|O|o|O|o|OE|oe|R|r|R|r|R|r|S"
    s = s & "|s|S|s|S|s|S|s|T|t|T|t|||U|u|U|u|U|u|U|u|U|u|U|u|W|w|Y|y|Y|Z|z|Z|z|Z|z|s|||||||||||||||||||||||||||||||||O|o||||||||||||||U|u||||||||||||||||||||DZ|Dz|dz|LJ|Lj|lj|NJ|Nj|nj|A|a|I|i|"
    s = s & "O|o|U|u|U|u|U|u|U|u|U|u||A|a|A|a|AE|ae|||G|g|K|k|O|o|O|o|||j|DZ|Dz|dz|G|g|||N|n|A|a|AE|ae|O|o|A|a|A|a|E|e|E|e|I|i|I|i|O|o|O|o|R|r|R|r|U|u|U|u|S|s|T|t|||H|h||||||Z|z|A|a|E|e|O|o|O|o|O|o|O|o|Y|y"
    s = s & "|||||||||||||||||||"
    partes = Split(s, "|")
    ReDim mMapaAscii(0 To 431)
    For i = 0 To 431
        If i <= UBound(partes) Then mMapaAscii(i) = partes(i) Else mMapaAscii(i) = ""
    Next i
    mMapaListo = True
End Sub

' Convierte cualquier valor de celda a texto (sin formato regional).
Public Function ValorTexto(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbEmpty, vbNull, vbError
            ValorTexto = ""
        Case vbString
            ValorTexto = Trim$(v)
        Case vbInteger, vbLong, vbByte
            ValorTexto = CStr(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal
            If v = Fix(v) And Abs(v) < 1E+15 Then
                ValorTexto = Format$(v, "0")
            Else
                ValorTexto = Replace(Trim$(Str$(v)), ",", ".")
            End If
        Case vbDate
            ValorTexto = FechaTexto(CDate(v))
        Case vbBoolean
            If v Then ValorTexto = "SI" Else ValorTexto = "NO"
        Case Else
            ValorTexto = Trim$(CStr(v))
    End Select
End Function

Public Function EsVacio(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbEmpty, vbNull
            EsVacio = True
        Case vbString
            EsVacio = (Trim$(v) = "")
        Case Else
            EsVacio = False
    End Select
End Function

' =====================================================================
'  IMPORTES (siempre Currency)
' =====================================================================

' Redondeo a 2 decimales "comercial" (0,005 -> 0,01), igual en cualquier PC.
Public Function Redondear2(ByVal v As Double) As Currency
    Dim x As Double
    x = Abs(v) * 100#
    x = Int(x + 0.5 + 0.000001)
    If v < 0 Then x = -x
    Redondear2 = CCur(x) / 100
End Function

' Importe con formato a3: signo + 10 enteros + punto + 2 decimales (14 posiciones).
Public Function ImporteA3(ByVal v As Currency) As String
    Dim signo As String, a As Currency, ent As Currency, cent As Long
    If v < 0 Then signo = "-" Else signo = "+"
    a = Abs(v)
    If a >= 1E+10 Then
        Err.Raise vbObjectError + 902, "ImporteA3", "Importe demasiado grande para a3: " & ImporteTexto(v)
    End If
    ent = Fix(a)
    cent = CLng((a - ent) * 100)
    If cent > 99 Then cent = 99
    ImporteA3 = signo & Right$("0000000000" & Format$(ent, "0"), 10) & "." & Right$("00" & CStr(cent), 2)
End Function

' Porcentaje con formato a3: xx.xx (5 posiciones).
Public Function PorcentajeA3(ByVal v As Currency) As String
    Dim a As Currency, ent As Currency, cent As Long
    a = Abs(v)
    If a >= 100 Then a = 99.99
    ent = Fix(a)
    cent = CLng((a - ent) * 100)
    If cent > 99 Then cent = 99
    PorcentajeA3 = Right$("00" & Format$(ent, "0"), 2) & "." & Right$("00" & CStr(cent), 2)
End Function

' Importe legible para mensajes: 1.234,56
Public Function ImporteTexto(ByVal v As Currency) As String
    Dim a As Currency, ent As Currency, cent As Long, s As String, grupos As String
    a = Abs(v)
    ent = Fix(a)
    cent = CLng((a - ent) * 100)
    s = Format$(ent, "0")
    grupos = ""
    Do While Len(s) > 3
        grupos = "." & Right$(s, 3) & grupos
        s = Left$(s, Len(s) - 3)
    Loop
    ImporteTexto = IIf(v < 0, "-", "") & s & grupos & "," & Right$("00" & CStr(cent), 2)
End Function

' Porcentaje legible: 10 / 5,5
Public Function PorcentajeTexto(ByVal v As Currency) As String
    Dim ent As Currency, cent As Long
    ent = Fix(Abs(v))
    cent = CLng((Abs(v) - ent) * 100)
    If cent = 0 Then
        PorcentajeTexto = IIf(v < 0, "-", "") & Format$(ent, "0")
    ElseIf cent Mod 10 = 0 Then
        PorcentajeTexto = IIf(v < 0, "-", "") & Format$(ent, "0") & "," & CStr(cent \ 10)
    Else
        PorcentajeTexto = IIf(v < 0, "-", "") & Format$(ent, "0") & "," & Right$("00" & CStr(cent), 2)
    End If
End Function

' Interpreta un número escrito como texto.
'   sepDecimal: "." , "," o "auto" (decide según el propio texto)
' Admite signo, espacios, símbolo €, paréntesis contables (12,50) y signo final 12,50-
Public Function ParseNumero(ByVal s As String, ByVal sepDecimal As String, ByRef ok As Boolean) As Double
    Dim negativo As Boolean, i As Long, ch As String, limpio As String
    Dim pPunto As Long, pComa As Long, dec As String, nPuntos As Long, nComas As Long
    ok = False
    s = Trim$(Replace(Replace(Replace(s, ChrW(160), ""), "EUR", "", , , vbTextCompare), ChrW(8364), ""))
    s = Replace(s, " ", "")
    If s = "" Then ok = True: ParseNumero = 0: Exit Function
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then negativo = True: s = Mid$(s, 2, Len(s) - 2)
    If Right$(s, 1) = "-" Then negativo = Not negativo: s = Left$(s, Len(s) - 1)
    If Left$(s, 1) = "-" Then
        negativo = Not negativo: s = Mid$(s, 2)
    ElseIf Left$(s, 1) = "+" Then
        s = Mid$(s, 2)
    End If
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            nPuntos = nPuntos + 1
        ElseIf ch = "," Then
            nComas = nComas + 1
        ElseIf ch < "0" Or ch > "9" Then
            Exit Function                         ' carácter no válido
        End If
    Next i
    pPunto = InStrRev(s, ".")
    pComa = InStrRev(s, ",")
    Select Case sepDecimal
        Case "."
            dec = "."
        Case ","
            dec = ","
        Case Else
            If nPuntos > 0 And nComas > 0 Then
                If pPunto > pComa Then dec = "." Else dec = ","
            ElseIf nComas > 0 Then
                If nComas = 1 Then dec = "," Else dec = "."
            ElseIf nPuntos = 1 Then
                ' "1.234" se lee como mil doscientos treinta y cuatro (costumbre española)
                If Len(s) - pPunto = 3 And Left$(s, 1) <> "0" Then dec = "," Else dec = "."
            Else
                dec = ","
            End If
    End Select
    If dec = "." Then
        If nPuntos > 1 Then Exit Function
        limpio = Replace(s, ",", "")
    Else
        If nComas > 1 Then Exit Function
        limpio = Replace(Replace(s, ".", ""), ",", ".")
    End If
    If limpio = "." Or limpio = "" Then Exit Function
    ParseNumero = Val(limpio)
    If negativo Then ParseNumero = -ParseNumero
    ok = True
End Function

' Lee un importe de una celda o de un campo CSV. ok = False si no es numérico.
Public Function LeerImporte(ByVal v As Variant, ByVal sepDecimal As String, ByRef ok As Boolean) As Currency
    Dim d As Double
    ok = True
    Select Case VarType(v)
        Case vbEmpty, vbNull
            LeerImporte = 0
        Case vbString
            d = ParseNumero(CStr(v), sepDecimal, ok)
            If ok Then
                If Abs(d) >= 10000000000# Then
                    ok = False                         ' a3 admite como máximo 10 enteros
                Else
                    LeerImporte = Redondear2(d)
                End If
            End If
        Case vbInteger, vbLong, vbByte, vbDouble, vbSingle, vbCurrency, vbDecimal
            d = CDbl(v)
            If Abs(d) >= 10000000000# Then
                ok = False
            Else
                LeerImporte = Redondear2(d)
            End If
        Case Else
            ok = False
    End Select
End Function

' Lee un porcentaje. Si la celda es numérica y está entre 0 y "limite" se
' entiende que tiene formato % de Excel (0,21 = 21 %) y se multiplica por 100
' ANTES de redondear (0,052 -> 5,20).
Public Function LeerPorcentaje(ByVal v As Variant, ByVal sepDecimal As String, ByVal limite As Double, ByRef ok As Boolean) As Currency
    Dim d As Double
    Select Case VarType(v)
        Case vbInteger, vbLong, vbByte, vbDouble, vbSingle, vbCurrency, vbDecimal
            d = CDbl(v)
            If d > 0 And d < limite Then
                ok = True
                LeerPorcentaje = Redondear2(d * 100)
                Exit Function
            End If
    End Select
    LeerPorcentaje = LeerImporte(v, sepDecimal, ok)
End Function

' =====================================================================
'  FECHAS
' =====================================================================
Public Function FechaA3(ByVal d As Date) As String
    FechaA3 = CStr(Year(d)) & Right$("0" & CStr(Month(d)), 2) & Right$("0" & CStr(Day(d)), 2)
End Function

Public Function FechaTexto(ByVal d As Date) As String
    FechaTexto = Right$("0" & CStr(Day(d)), 2) & "/" & Right$("0" & CStr(Month(d)), 2) & "/" & CStr(Year(d))
End Function

Public Function MarcaTiempo() As String
    Dim t As Date
    t = Now
    MarcaTiempo = FechaA3(t) & "_" & Right$("0" & CStr(Hour(t)), 2) & Right$("0" & CStr(Minute(t)), 2) & _
                  Right$("0" & CStr(Second(t)), 2)
End Function

' Lee una fecha de una celda o de un texto.
'   formato: "DMA" (dd/mm/aaaa, por defecto), "MDA" (mm/dd/aaaa) o "AMD"
Public Function LeerFecha(ByVal v As Variant, ByVal formato As String, ByRef ok As Boolean) As Date
    Dim d As Double, res As Date
    ok = False
    Select Case VarType(v)
        Case vbDate
            LeerFecha = DateSerial(Year(v), Month(v), Day(v))
            ok = True
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency, vbDecimal
            d = CDbl(v)
            If d >= 20000101 And d <= 20991231 And d = Fix(d) Then
                ok = FechaDesdePartes(CLng(Fix(d / 10000)), CLng(Fix(d / 100)) Mod 100, CLng(d) Mod 100, res)
                If ok Then LeerFecha = res
            ElseIf d >= 1 And d < 2958466 Then
                LeerFecha = DateSerial(1899, 12, 30) + Int(d)
                ok = True
            End If
        Case vbString
            LeerFecha = ParseFechaTexto(CStr(v), formato, ok)
    End Select
End Function

Public Function ParseFechaTexto(ByVal s As String, ByVal formato As String, ByRef ok As Boolean) As Date
    Dim p() As String, a As Long, b As Long, c As Long, i As Long, t As String, res As Date
    ok = False
    s = Trim$(s)
    If s = "" Then Exit Function
    ' quitar la hora si la hay ("22/09/2026 10:15" o "2026-09-22T10:15:00")
    i = InStr(s, "T")
    If i > 8 Then s = Left$(s, i - 1)
    i = InStr(s, " ")
    If i > 0 Then s = Left$(s, i - 1)
    ' sólo dígitos: aaaammdd o ddmmaaaa
    If Len(s) = 8 And SoloDigitos(s) Then
        If Left$(s, 2) = "19" Or Left$(s, 2) = "20" Then
            ok = FechaDesdePartes(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Right$(s, 2)), res)
        End If
        If Not ok Then ok = FechaDesdePartes(CLng(Right$(s, 4)), CLng(Mid$(s, 3, 2)), CLng(Left$(s, 2)), res)
        If ok Then ParseFechaTexto = res
        Exit Function
    End If
    t = Replace(Replace(s, "-", "/"), ".", "/")
    p = Split(t, "/")
    If UBound(p) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(p(i)) = 0 Or Len(p(i)) > 4 Then Exit Function
        If Not SoloDigitos(p(i)) Then Exit Function
    Next i
    If Len(p(0)) = 4 Or UCase$(formato) = "AMD" Then
        a = CLng(p(0)): b = CLng(p(1)): c = CLng(p(2))           ' año, mes, día
    ElseIf UCase$(formato) = "MDA" Then
        a = CLng(p(2)): b = CLng(p(0)): c = CLng(p(1))
    Else
        a = CLng(p(2)): b = CLng(p(1)): c = CLng(p(0))
    End If
    If a < 100 Then a = a + 2000
    ok = FechaDesdePartes(a, b, c, res)
    If ok Then ParseFechaTexto = res
End Function

Private Function FechaDesdePartes(ByVal anio As Long, ByVal mes As Long, ByVal dia As Long, ByRef resultado As Date) As Boolean
    Dim d As Date
    If anio < 1990 Or anio > 2099 Or mes < 1 Or mes > 12 Or dia < 1 Or dia > 31 Then Exit Function
    d = DateSerial(anio, mes, dia)
    If Month(d) <> mes Or Day(d) <> dia Then Exit Function      ' p.ej. 31/02
    resultado = d
    FechaDesdePartes = True
End Function

Public Function SoloDigitos(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    SoloDigitos = True
End Function

' =====================================================================
'  CUENTAS Y DOCUMENTOS
' =====================================================================

' Devuelve la cuenta con la longitud del plan. Admite la notación de
' punto de a3 ("430.25" con 8 dígitos -> "43000025").
' Si hay algún problema devuelve "" y explica el motivo en "motivo".
Public Function NormalizarCuenta(ByVal cuenta As String, ByVal digitos As Integer, ByRef motivo As String) As String
    Dim p As Long, izq As String, der As String, relleno As Long
    motivo = ""
    cuenta = Replace(Replace(Trim$(cuenta), " ", ""), ChrW(160), "")
    If cuenta = "" Then
        motivo = "cuenta vacía"
        Exit Function
    End If
    p = InStr(cuenta, ".")
    If p > 0 Then
        If InStr(p + 1, cuenta, ".") > 0 Then
            motivo = "la cuenta " & cuenta & " tiene más de un punto"
            Exit Function
        End If
        izq = Left$(cuenta, p - 1)
        der = Mid$(cuenta, p + 1)
        relleno = digitos - Len(izq) - Len(der)
        If relleno < 0 Or izq = "" Then
            motivo = "la cuenta " & cuenta & " no cabe en " & digitos & " dígitos"
            Exit Function
        End If
        cuenta = izq & String$(relleno, "0") & der
    End If
    If Not SoloDigitos(cuenta) Then
        motivo = "la cuenta " & cuenta & " contiene caracteres que no son números"
        Exit Function
    End If
    If Len(cuenta) <> digitos Then
        motivo = "la cuenta " & cuenta & " tiene " & Len(cuenta) & " dígitos y el plan de la empresa tiene " & digitos
        Exit Function
    End If
    NormalizarCuenta = cuenta
End Function

' Cuenta por defecto a partir de un prefijo: CuentaPatron("477", 8, "21") -> "47700021"
Public Function CuentaPatron(ByVal prefijo As String, ByVal digitos As Integer, ByVal sufijo As String) As String
    Dim n As Long
    n = digitos - Len(prefijo) - Len(sufijo)
    If n < 0 Then n = 0
    CuentaPatron = Left$(prefijo & String$(n, "0") & sufijo, digitos)
End Function

' Serie de un documento: "INV347" -> "INV", "CRN-8" -> "CRN", "2026/0015" -> "2026"
Public Function SerieDocumento(ByVal doc As String) As String
    Dim i As Long
    doc = Trim$(doc)
    i = Len(doc)
    Do While i > 0
        If Mid$(doc, i, 1) < "0" Or Mid$(doc, i, 1) > "9" Then Exit Do
        i = i - 1
    Loop
    SerieDocumento = Normalizar(Left$(doc, i))
End Function

' Número final de un documento: "INV347" -> 347 ; sin número -> -1
Public Function NumeroDocumento(ByVal doc As String) As Double
    Dim i As Long
    doc = Trim$(doc)
    i = Len(doc)
    Do While i > 0
        If Mid$(doc, i, 1) < "0" Or Mid$(doc, i, 1) > "9" Then Exit Do
        i = i - 1
    Loop
    If i = Len(doc) Then
        NumeroDocumento = -1
    Else
        NumeroDocumento = Val(Mid$(doc, i + 1))
    End If
End Function

' ¿Pasa el documento el filtro de rango (misma serie y número entre desde y hasta)?
Public Function DocumentoEnRango(ByVal doc As String, ByVal desde As String, ByVal hasta As String) As Boolean
    Dim serie As String, num As Double, ref As String
    desde = Trim$(desde): hasta = Trim$(hasta)
    If desde = "" And hasta = "" Then DocumentoEnRango = True: Exit Function
    If desde <> "" Then ref = SerieDocumento(desde) Else ref = SerieDocumento(hasta)
    serie = SerieDocumento(doc)
    num = NumeroDocumento(doc)
    If serie <> ref Or num < 0 Then Exit Function
    If desde <> "" Then
        If NumeroDocumento(desde) >= 0 And num < NumeroDocumento(desde) Then Exit Function
    End If
    If hasta <> "" Then
        If NumeroDocumento(hasta) >= 0 And num > NumeroDocumento(hasta) Then Exit Function
    End If
    DocumentoEnRango = True
End Function

' ¿Pasa la serie los filtros de series a incluir / excluir?
Public Function SerieAdmitida(ByVal doc As String, ByVal incluir As String, ByVal excluir As String) As Boolean
    Dim serie As String
    serie = SerieDocumento(doc)
    If Trim$(incluir) <> "" Then
        If Not SerieEnLista(serie, incluir) Then Exit Function
    End If
    If Trim$(excluir) <> "" Then
        If SerieEnLista(serie, excluir) Then Exit Function
    End If
    SerieAdmitida = True
End Function

Private Function SerieEnLista(ByVal serie As String, ByVal lista As String) As Boolean
    Dim partes() As String, i As Long
    partes = Split(lista, ";")
    For i = LBound(partes) To UBound(partes)
        If SerieCoincide(serie, partes(i)) Then
            SerieEnLista = True
            Exit Function
        End If
    Next i
End Function

' La serie escrita coincide si es la serie completa o su comienzo hasta un
' separador: "F" vale para "F-2026-001" (serie "F 2026") y "FV" para
' "FV2026-000017" (serie "FV2026"), pero "F" no vale para "FV2026".
Private Function SerieCoincide(ByVal serie As String, ByVal valor As String) As Boolean
    Dim sig As String, ult As String
    valor = Normalizar(valor)
    If valor = "" Or serie = "" Then Exit Function
    If Left$(serie, Len(valor)) <> valor Then Exit Function
    If Len(serie) = Len(valor) Then
        SerieCoincide = True
        Exit Function
    End If
    sig = Mid$(serie, Len(valor) + 1, 1)
    ult = Right$(valor, 1)
    If sig = " " Then
        SerieCoincide = True
    ElseIf sig >= "0" And sig <= "9" Then
        SerieCoincide = Not (ult >= "0" And ult <= "9")
    End If
End Function

' Número de factura tal como cabe en a3 (10 posiciones). Si es más largo se
' quitan separadores y se conserva la parte final, que es la que distingue
' una factura de otra (F2026/000123 -> 2026000123).
Public Function DocumentoA3(ByVal doc As String) As String
    doc = AsciiA3(doc)
    If Len(doc) > 10 Then doc = Replace(Replace(Replace(Replace(doc, "-", ""), "/", ""), " ", ""), ".", "")
    If Len(doc) > 10 Then doc = Right$(doc, 10)
    DocumentoA3 = doc
End Function

' Normaliza un NIF: mayúsculas, sin espacios ni separadores y sin prefijo ES.
Public Function NormalizarNIF(ByVal nif As String) As String
    Dim s As String
    s = UCase$(AsciiA3(nif))
    s = Replace(Replace(Replace(Replace(Replace(s, " ", ""), "-", ""), ".", ""), "/", ""), "_", "")
    If Len(s) = 11 And Left$(s, 2) = "ES" Then
        If NIFValido(Mid$(s, 3)) Then s = Mid$(s, 3)
    End If
    NormalizarNIF = s
End Function

' Comprueba la letra o el dígito de control de un DNI, NIE o CIF español.
Public Function NIFValido(ByVal nif As String) As Boolean
    Const LETRAS As String = "TRWAGMYFPDXBNJZSQVHLCKE"
    Dim cuerpo As String, c As String, i As Long, sumaPar As Long, sumaImpar As Long, d As Long, control As Long
    If Len(nif) <> 9 Then Exit Function
    c = Left$(nif, 1)
    If c = "X" Or c = "Y" Or c = "Z" Then
        cuerpo = CStr(InStr("XYZ", c) - 1) & Mid$(nif, 2, 7)
        If Not SoloDigitos(cuerpo) Then Exit Function
        NIFValido = (Mid$(LETRAS, (CLng(cuerpo) Mod 23) + 1, 1) = Right$(nif, 1))
    ElseIf SoloDigitos(Left$(nif, 8)) Then
        NIFValido = (Mid$(LETRAS, (CLng(Left$(nif, 8)) Mod 23) + 1, 1) = Right$(nif, 1))
    ElseIf InStr("ABCDEFGHJKLMNPQRSUVW", c) > 0 And SoloDigitos(Mid$(nif, 2, 7)) Then
        For i = 2 To 8
            d = CLng(Mid$(nif, i, 1))
            If i Mod 2 = 1 Then
                sumaPar = sumaPar + d                          ' posiciones 2, 4 y 6 del número
            Else
                d = d * 2
                sumaImpar = sumaImpar + (d \ 10) + (d Mod 10)
            End If
        Next i
        control = (10 - ((sumaPar + sumaImpar) Mod 10)) Mod 10
        c = Right$(nif, 1)
        NIFValido = (c = CStr(control) Or c = Mid$("JABCDEFGHI", control + 1, 1))
    End If
End Function

' =====================================================================
'  CONSTRUCCIÓN DE REGISTROS DEL SUENLACE.DAT (formato 4, 254 posiciones)
' =====================================================================

' Registro 0: apunte sin IVA (diarios)
Public Function RegistroApunte(ByVal empresa As String, ByVal fecha As Date, ByVal cuenta As String, _
        ByVal descCuenta As String, ByVal debeHaber As String, ByVal documento As String, _
        ByVal lineaApunte As String, ByVal descApunte As String, ByVal importe As Currency) As String
    Dim r As String
    r = "4" & CodigoEmpresaA3(empresa) & FechaA3(fecha) & "0" & _
        TextoA3(cuenta, 12) & TextoA3(descCuenta, 30) & debeHaber & TextoA3(documento, 10) & _
        lineaApunte & TextoA3(descApunte, 30) & ImporteA3(importe) & _
        Space$(138) & " " & "E" & "N"
    RegistroApunte = r
End Function

' Registro 1 (factura) / 2 (abono): cabecera de factura con IVA
Public Function RegistroCabeceraFactura(ByVal empresa As String, ByVal fecha As Date, ByVal esAbono As Boolean, _
        ByVal cuenta As String, ByVal descCuenta As String, ByVal tipoFactura As String, _
        ByVal documento As String, ByVal descApunte As String, ByVal total As Currency, _
        ByVal nif As String, ByVal nombre As String, ByVal cp As String, _
        ByVal fechaOperacion As Date, ByVal fechaFactura As Date) As String
    Dim r As String
    r = "4" & CodigoEmpresaA3(empresa) & FechaA3(fecha) & IIf(esAbono, "2", "1") & _
        TextoA3(cuenta, 12) & TextoA3(descCuenta, 30) & tipoFactura & TextoA3(DocumentoA3(documento), 10) & _
        "I" & TextoA3(descApunte, 30) & ImporteA3(total) & Space$(62) & _
        TextoA3(nif, 14) & TextoA3(nombre, 40) & TextoA3(cp, 5) & "  " & _
        FechaA3(fechaOperacion) & FechaA3(fechaFactura) & "E" & "N"
    RegistroCabeceraFactura = r
End Function

' Registro 9: detalle de IVA (una línea por cuenta de ventas y tipo de IVA)
'   tipoImporte: "C" cargo (lo normal) / "A" abono en factura
'   marcaCaja (177) y marca0 (178): en blanco salvo casos especiales
'   ctaIVA2 / ctaRE2 (228-251): IVA y recargo repercutido de la autorrepercusión
'   (facturas recibidas con inversión del sujeto pasivo o intracomunitarias)
Public Function RegistroDetalleIVA(ByVal empresa As String, ByVal fecha As Date, ByVal cuenta As String, _
        ByVal descCuenta As String, ByVal tipoImporte As String, ByVal documento As String, ByVal lineaApunte As String, _
        ByVal descApunte As String, ByVal subtipo As String, ByVal baseImp As Currency, _
        ByVal pctIVA As Currency, ByVal cuota As Currency, ByVal pctRE As Currency, ByVal cuotaRE As Currency, _
        ByVal pctRet As Currency, ByVal cuotaRet As Currency, ByVal impreso As String, _
        ByVal marcaCaja As String, ByVal marca0 As String, _
        ByVal ctaIVA As String, ByVal ctaRE As String, ByVal ctaRet As String, _
        ByVal ctaIVA2 As String, ByVal ctaRE2 As String) As String
    Dim r As String
    r = "4" & CodigoEmpresaA3(empresa) & FechaA3(fecha) & "9" & _
        TextoA3(cuenta, 12) & TextoA3(descCuenta, 30) & Left$(tipoImporte & "C", 1) & _
        TextoA3(DocumentoA3(documento), 10) & _
        lineaApunte & TextoA3(descApunte, 30) & Right$("00" & subtipo, 2) & _
        ImporteA3(baseImp) & PorcentajeA3(pctIVA) & ImporteA3(cuota) & _
        PorcentajeA3(pctRE) & ImporteA3(cuotaRE) & PorcentajeA3(pctRet) & ImporteA3(cuotaRet) & _
        Right$("00" & impreso, 2) & "S" & "N" & Left$(marcaCaja & " ", 1) & Left$(marca0 & " ", 1) & Space$(13) & _
        TextoA3(ctaIVA, 12) & TextoA3(ctaRE, 12) & TextoA3(ctaRet, 12) & TextoA3(ctaIVA2, 12) & TextoA3(ctaRE2, 12) & _
        " " & "E" & "N"
    RegistroDetalleIVA = r
End Function

Public Function CodigoEmpresaA3(ByVal codigo As String) As String
    CodigoEmpresaA3 = Right$("00000" & Trim$(codigo), 5)
End Function

' =====================================================================
'  FICHEROS
' =====================================================================
Public Function ExisteFichero(ByVal ruta As String) As Boolean
    On Error Resume Next
    ExisteFichero = (Len(Dir$(ruta, vbNormal Or vbReadOnly Or vbHidden)) > 0)
    If Err.Number <> 0 Then ExisteFichero = False
    On Error GoTo 0
End Function

Public Function ExisteCarpeta(ByVal ruta As String) As Boolean
    On Error Resume Next
    If Right$(ruta, 1) = "\" Or Right$(ruta, 1) = "/" Then ruta = Left$(ruta, Len(ruta) - 1)
    ExisteCarpeta = ((GetAttr(ruta) And vbDirectory) = vbDirectory)
    If Err.Number <> 0 Then ExisteCarpeta = False
    On Error GoTo 0
End Function

' Crea la carpeta (y las intermedias) si no existe.
Public Function CrearCarpeta(ByVal ruta As String) As Boolean
    Dim partes() As String, i As Long, actual As String, sep As String
    If ExisteCarpeta(ruta) Then CrearCarpeta = True: Exit Function
    sep = SeparadorRuta()
    ruta = Replace(ruta, "/", sep)
    ruta = Replace(ruta, "\", sep)
    partes = Split(ruta, sep)
    On Error Resume Next
    For i = LBound(partes) To UBound(partes)
        If i = LBound(partes) Then
            actual = partes(i)
        Else
            actual = actual & sep & partes(i)
        End If
        If Len(actual) > 0 And Right$(actual, 1) <> ":" Then
            If Not ExisteCarpeta(actual) Then MkDir actual
        End If
    Next i
    On Error GoTo 0
    CrearCarpeta = ExisteCarpeta(ruta)
End Function

Public Function SeparadorRuta() As String
    If InStr(CurDir$, "/") > 0 And InStr(CurDir$, "\") = 0 Then SeparadorRuta = "/" Else SeparadorRuta = "\"
End Function

Public Function UnirRuta(ByVal carpeta As String, ByVal fichero As String) As String
    Dim sep As String
    sep = SeparadorRuta()
    If Right$(carpeta, 1) = "\" Or Right$(carpeta, 1) = "/" Then
        UnirRuta = carpeta & fichero
    Else
        UnirRuta = carpeta & sep & fichero
    End If
End Function

Public Function CarpetaDeRuta(ByVal ruta As String) As String
    Dim p As Long
    p = InStrRev(ruta, "\")
    If InStrRev(ruta, "/") > p Then p = InStrRev(ruta, "/")
    If p > 0 Then CarpetaDeRuta = Left$(ruta, p - 1)
End Function

Public Function NombreDeRuta(ByVal ruta As String) As String
    Dim p As Long
    p = InStrRev(ruta, "\")
    If InStrRev(ruta, "/") > p Then p = InStrRev(ruta, "/")
    NombreDeRuta = Mid$(ruta, p + 1)
End Function

' Si ya existe un SUENLACE.DAT lo renombra a SUENLACE_anterior_aaaammdd_hhmmss.DAT
Public Function CopiaSeguridadDat(ByVal ruta As String) As String
    Dim destino As String
    If Not ExisteFichero(ruta) Then Exit Function
    destino = UnirRuta(CarpetaDeRuta(ruta), "SUENLACE_anterior_" & MarcaTiempo() & ".DAT")
    Name ruta As destino
    CopiaSeguridadDat = destino
End Function

' Escribe las líneas en disco: ASCII puro, cada línea + CR/LF.
Public Sub EscribirDat(ByVal ruta As String)
    Dim b() As Byte, i As Long, j As Long, k As Long, f As Integer, c As Long, linea As String
    If gNDat = 0 Then Err.Raise vbObjectError + 903, "EscribirDat", "No hay ninguna línea que escribir."
    ReDim b(0 To gNDat * (A3_LONG_REGISTRO + 2) - 1)
    k = 0
    For i = 1 To gNDat
        linea = gDat(i)
        For j = 1 To Len(linea)
            c = AscW(Mid$(linea, j, 1))
            If c < 32 Or c > 126 Then c = 32              ' seguridad: nunca un carácter no ASCII
            b(k) = c
            k = k + 1
        Next j
        b(k) = 13: b(k + 1) = 10
        k = k + 2
    Next i
    If ExisteFichero(ruta) Then Kill ruta
    f = FreeFile
    Open ruta For Binary Access Write As #f
    Put #f, , b
    Close #f
End Sub

' Lee un fichero de texto detectando la codificación (UTF-8 con o sin BOM,
' UTF-16 o ANSI Windows-1252).
Public Function LeerFicheroTexto(ByVal ruta As String) As String
    Dim b() As Byte, f As Integer, n As Long, ok As Boolean, s As String
    f = FreeFile
    Open ruta For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then
        Close #f
        LeerFicheroTexto = ""
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, , b
    Close #f
    If n >= 2 Then
        If b(0) = &HFF And b(1) = &HFE Then
            LeerFicheroTexto = DecodificarUTF16(b, 2)
            Exit Function
        End If
    End If
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
            LeerFicheroTexto = DecodificarUTF8(b, 3, ok)
            If ok Then Exit Function
        End If
    End If
    s = DecodificarUTF8(b, 0, ok)
    If ok Then
        LeerFicheroTexto = s
    Else
        LeerFicheroTexto = DecodificarANSI(b)
    End If
End Function

Private Function DecodificarUTF8(ByRef b() As Byte, ByVal inicio As Long, ByRef ok As Boolean) As String
    Dim i As Long, n As Long, c As Long, c2 As Long, c3 As Long, c4 As Long, cp As Long
    Dim buf As String, pos As Long
    ok = False
    n = UBound(b)
    buf = Space$(n - inicio + 2)
    pos = 0
    i = inicio
    Do While i <= n
        c = b(i)
        If c < &H80 Then
            cp = c: i = i + 1
        ElseIf c >= &HC2 And c <= &HDF Then
            If i + 1 > n Then Exit Function
            c2 = b(i + 1)
            If (c2 And &HC0) <> &H80 Then Exit Function
            cp = ((c And &H1F) * 64) Or (c2 And &H3F): i = i + 2
        ElseIf c >= &HE0 And c <= &HEF Then
            If i + 2 > n Then Exit Function
            c2 = b(i + 1): c3 = b(i + 2)
            If (c2 And &HC0) <> &H80 Or (c3 And &HC0) <> &H80 Then Exit Function
            cp = ((c And &HF) * 4096) Or ((c2 And &H3F) * 64) Or (c3 And &H3F): i = i + 3
            If cp < &H800 Then Exit Function
        ElseIf c >= &HF0 And c <= &HF4 Then
            If i + 3 > n Then Exit Function
            c2 = b(i + 1): c3 = b(i + 2): c4 = b(i + 3)
            If (c2 And &HC0) <> &H80 Or (c3 And &HC0) <> &H80 Or (c4 And &HC0) <> &H80 Then Exit Function
            cp = &HFFFD: i = i + 4                   ' fuera del plano básico: se descarta
        Else
            Exit Function
        End If
        pos = pos + 1
        Mid$(buf, pos, 1) = ChrW(cp)
    Loop
    DecodificarUTF8 = Left$(buf, pos)
    ok = True
End Function

Private Function DecodificarUTF16(ByRef b() As Byte, ByVal inicio As Long) As String
    Dim i As Long, buf As String, pos As Long, cp As Long
    buf = Space$((UBound(b) - inicio + 2) \ 2)
    For i = inicio To UBound(b) - 1 Step 2
        cp = b(i) + b(i + 1) * 256&
        pos = pos + 1
        If cp >= &HD800& And cp <= &HDFFF& Then cp = &HFFFD
        Mid$(buf, pos, 1) = ChrW(cp)
    Next i
    DecodificarUTF16 = Left$(buf, pos)
End Function

Private Function DecodificarANSI(ByRef b() As Byte) As String
    ' Windows-1252: 0x80-0x9F tienen caracteres propios; el resto coincide con Latin-1
    Dim tabla As Variant, i As Long, buf As String, cp As Long
    tabla = Array(8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141, 381, 143, _
                  144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250, 339, 157, 382, 376)
    buf = Space$(UBound(b) + 1)
    For i = 0 To UBound(b)
        cp = b(i)
        If cp >= &H80 And cp <= &H9F Then cp = tabla(cp - &H80)
        Mid$(buf, i + 1, 1) = ChrW(cp)
    Next i
    DecodificarANSI = buf
End Function

' Convierte un texto CSV en una matriz (1..filas, 1..columnas) de textos.
'   separador: ";", ",", vbTab, "|"  o  "" / "auto" para detectarlo en la primera línea.
Public Function ParsearCSV(ByVal texto As String, ByVal separador As String) As Variant
    Dim n As Long, i As Long, ch As String, enComillas As Boolean
    Dim campo As String, campos() As String, nCampos As Long, filaDe() As Long
    Dim fila As Long, maxCol As Long, colActual As Long, resultado() As Variant
    Dim r As Long, c As Long, k As Long, sep As String, finFila As Boolean

    sep = SeparadorCSV(texto, separador)
    n = Len(texto)
    ReDim campos(1 To 1024)
    ReDim filaDe(1 To 1024)
    fila = 1: colActual = 0: nCampos = 0: campo = ""
    i = 1
    Do While i <= n
        ch = Mid$(texto, i, 1)
        finFila = False
        If enComillas Then
            If ch = """" Then
                If i < n And Mid$(texto, i + 1, 1) = """" Then
                    campo = campo & """": i = i + 1
                Else
                    enComillas = False
                End If
            Else
                campo = campo & ch
            End If
        ElseIf ch = """" Then
            enComillas = True
        ElseIf ch = sep Then
            GoSub GuardarCampo
        ElseIf ch = vbCr Or ch = vbLf Then
            If ch = vbCr And i < n Then
                If Mid$(texto, i + 1, 1) = vbLf Then i = i + 1
            End If
            GoSub GuardarCampo
            If colActual > maxCol Then maxCol = colActual
            fila = fila + 1: colActual = 0
            finFila = True
        Else
            campo = campo & ch
        End If
        i = i + 1
    Loop
    If Not finFila Then
        GoSub GuardarCampo
        If colActual > maxCol Then maxCol = colActual
    Else
        fila = fila - 1
    End If
    ' quitar filas finales vacías
    Do While fila > 0
        If Not FilaCSVVacia(campos, filaDe, nCampos, fila) Then Exit Do
        fila = fila - 1
    Loop
    If fila < 1 Or maxCol < 1 Then
        ReDim resultado(1 To 1, 1 To 1)
        ParsearCSV = resultado
        Exit Function
    End If
    ReDim resultado(1 To fila, 1 To maxCol)
    r = 0: c = 0
    For k = 1 To nCampos
        If filaDe(k) <> r Then r = filaDe(k): c = 0
        c = c + 1
        If r <= fila Then resultado(r, c) = campos(k)
    Next k
    ParsearCSV = resultado
    Exit Function

GuardarCampo:
    nCampos = nCampos + 1
    If nCampos > UBound(campos) Then
        ReDim Preserve campos(1 To UBound(campos) * 2)
        ReDim Preserve filaDe(1 To UBound(filaDe) * 2)
    End If
    campos(nCampos) = campo
    filaDe(nCampos) = fila
    colActual = colActual + 1
    campo = ""
    Return
End Function

Private Function FilaCSVVacia(ByRef campos() As String, ByRef filaDe() As Long, ByVal nCampos As Long, ByVal fila As Long) As Boolean
    Dim k As Long
    For k = nCampos To 1 Step -1
        If filaDe(k) < fila Then Exit For
        If filaDe(k) = fila Then
            If Trim$(campos(k)) <> "" Then Exit Function
        End If
    Next k
    FilaCSVVacia = True
End Function

' Detecta el separador contando ; , TAB y | fuera de comillas en la primera línea.
Public Function SeparadorCSV(ByVal texto As String, ByVal separador As String) As String
    Dim i As Long, ch As String, enComillas As Boolean
    Dim nPC As Long, nC As Long, nT As Long, nB As Long
    Select Case UCase$(Trim$(separador))
        Case ";", ",", "|"
            SeparadorCSV = Trim$(separador): Exit Function
        Case "TAB", vbTab
            SeparadorCSV = vbTab: Exit Function
    End Select
    For i = 1 To Len(texto)
        ch = Mid$(texto, i, 1)
        If ch = """" Then
            enComillas = Not enComillas
        ElseIf Not enComillas Then
            Select Case ch
                Case ";": nPC = nPC + 1
                Case ",": nC = nC + 1
                Case vbTab: nT = nT + 1
                Case "|": nB = nB + 1
                Case vbCr, vbLf: Exit For
            End Select
        End If
    Next i
    SeparadorCSV = ";"
    If nC > nPC And nC >= nT And nC >= nB Then SeparadorCSV = ","
    If nT > nPC And nT > nC And nT >= nB Then SeparadorCSV = vbTab
    If nB > nPC And nB > nC And nB > nT Then SeparadorCSV = "|"
End Function

' Lee un elemento de una colección. Devuelve False si la clave no existe.
Public Function LeerColeccion(ByVal c As Collection, ByVal clave As String, ByRef valor As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c(clave)
    If Err.Number = 0 Then
        valor = v
        LeerColeccion = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Pasa cualquier valor de Range.Value a matriz 2D (1..n, 1..m)
Public Function ComoMatriz(ByVal v As Variant) As Variant
    Dim m() As Variant
    If IsArray(v) Then
        ComoMatriz = v
    Else
        ReDim m(1 To 1, 1 To 1)
        m(1, 1) = v
        ComoMatriz = m
    End If
End Function

' Número de columnas de una matriz 2D (0 si no es matriz)
Public Function ColumnasMatriz(ByRef m As Variant) As Long
    On Error Resume Next
    ColumnasMatriz = UBound(m, 2)
    If Err.Number <> 0 Then ColumnasMatriz = 0
    On Error GoTo 0
End Function

' ¿Está vacía toda la fila r de la matriz?
Public Function FilaVacia(ByRef m As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = LBound(m, 2) To UBound(m, 2)
        If Not EsVacio(m(r, c)) Then Exit Function
    Next c
    FilaVacia = True
End Function

' Busca una columna por su cabecera. "nombres" admite alternativas separadas por "|".
' Devuelve 0 si no la encuentra.
Public Function BuscarColumna(ByRef m As Variant, ByVal filaCab As Long, ByVal nombres As String) As Long
    Dim alt() As String, i As Long, c As Long, buscado As String
    If Trim$(nombres) = "" Then Exit Function
    alt = Split(nombres, "|")
    For i = LBound(alt) To UBound(alt)
        buscado = Normalizar(alt(i))
        If buscado <> "" Then
            For c = LBound(m, 2) To UBound(m, 2)
                If Normalizar(ValorTexto(m(filaCab, c))) = buscado Then
                    BuscarColumna = c
                    Exit Function
                End If
            Next c
        End If
    Next i
End Function
