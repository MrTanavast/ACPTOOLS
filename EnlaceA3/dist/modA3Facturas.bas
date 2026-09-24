Attribute VB_Name = "modA3Facturas"
' =====================================================================
'  ENLACE CONTABLE A3  ·  ACP ASOCIADOS
'  Módulo LIBRO DE FACTURAS EMITIDAS
'
'  Lee el libro (CSV o Excel) según el perfil, agrupa las líneas por
'  documento (nº de factura + fecha), aplica las reglas de exclusión y
'  genera por cada factura:
'     - 1 registro de cabecera (tipo 1 = factura, tipo 2 = abono)
'     - 1 registro de detalle (tipo 9) por cuenta de ventas y tipo de IVA
'
'  Reglas (las mismas que el procedimiento de facturas de ventas):
'     - Documentos anulados            -> excluidos
'     - Importes no numéricos          -> excluidos (revisar)
'     - Documentos con total 0         -> excluidos
'     - Líneas con base y cuota a 0    -> se quitan de la factura
'     - Tipo de IVA no configurado     -> excluido (contabilizar a mano)
'     - Total distinto de base+cuotas  -> excluido (a3 lo rechazaría)
'     - Abonos: registro tipo 2 con importes en positivo
'     - Tickets: se tratan como facturas
' =====================================================================
Option Explicit

Public Const TIPO_FACTURA As Integer = 1
Public Const TIPO_TICKET As Integer = 2
Public Const TIPO_ABONO As Integer = 3
Public Const TIPO_RECTIFICATIVA As Integer = 4

' Línea del libro de facturas ya interpretada
Public Type TLinFac
    Fila As Long
    FechaOk As Boolean
    Fecha As Date
    ClaveFecha As String
    FechaOp As Date
    Documento As String
    Tipo As Integer
    TipoTexto As String
    Anulada As Boolean
    Cliente As String
    NIF As String
    CP As String
    CtaCliente As String
    DescCtaCliente As String
    CtaVentas As String
    DescCtaVentas As String
    CtaIVA As String
    CtaRE As String
    CtaRet As String
    BaseImp As Currency
    PctIVA As Currency
    Cuota As Currency
    PctRE As Currency
    CuotaRE As Currency
    PctRet As Currency
    CuotaRet As Currency
    Total As Currency
    Subtipo As String
    Impreso As String
    ErrorLectura As String
    AvisoLectura As String
    DentroFiltro As Boolean
    Doc As Long
    Sig As Long
    Estado As String
End Type

' Documento (factura) = grupo de líneas
Public Type TDocFac
    Documento As String
    Fecha As Date
    Primera As Long
    Ultima As Long
    NLineas As Long
    ClaveBase As String
    Anulado As Boolean
    ConAnulado As Boolean
End Type

' Línea de control: lo que va exactamente al SUENLACE (abonos en negativo)
Public Type TCtrlFac
    Fecha As Date
    Documento As String
    TipoTexto As String
    Cliente As String
    NIF As String
    CtaCliente As String
    CtaVentas As String
    PctIVA As Currency
    CtaIVA As String
    BaseImp As Currency
    Cuota As Currency
    CuotaRE As Currency
    CuotaRet As Currency
    Total As Currency
End Type

' Detalle agregado de una factura (cuenta de ventas + tipo de IVA)
Private Type TDetFac
    CtaVentas As String
    DescVentas As String
    PctIVA As Currency
    CtaIVA As String
    PctRE As Currency
    CtaRE As String
    PctRet As Currency
    CtaRet As String
    Subtipo As String
    Impreso As String
    BaseImp As Currency
    Cuota As Currency
    CuotaRE As Currency
    CuotaRet As Currency
End Type

Public gLinFac() As TLinFac
Public gNLinFac As Long
Public gDocFac() As TDocFac
Public gNDocFac As Long
Public gCtrlFac() As TCtrlFac
Public gNCtrlFac As Long
Private mNumerosA3 As Collection

' =====================================================================
'  PROCESO COMPLETO
' =====================================================================
Public Function FacturasProcesar(ByRef emp As TEmpresa, ByRef per As TPerfil, ByRef fil As TFiltros, _
                                 ByRef datos As Variant, ByRef res As TResumen, ByRef msgError As String) As Boolean
    Dim col(1 To NUM_CAMPOS) As Long, d As Long

    ResumenReiniciar res
    msgError = ""
    DatReiniciar
    IncReiniciar
    gNLinFac = 0: ReDim gLinFac(1 To 16)
    gNDocFac = 0: ReDim gDocFac(1 To 16)
    gNCtrlFac = 0: ReDim gCtrlFac(1 To 16)
    Set mNumerosA3 = New Collection

    If Not ResolverColumnasFacturas(per, datos, col, msgError) Then Exit Function
    LeerLineasFacturas emp, per, datos, col
    AplicarFiltrosFacturas fil, res
    AgruparDocumentosFacturas

    For d = 1 To gNDocFac
        If TipoAdmitido(fil, TipoDocumento(per, d)) Then
            ProcesarDocumentoFactura emp, per, d, res
        Else
            FiltrarDocumento d, res
        End If
    Next d

    res.LineasDat = gNDat
    res.Avisos = IncContar(INC_AVISO)
    res.Excluidos = IncContar(INC_EXCLUIDO)
    FacturasProcesar = True
End Function

' Localiza cada campo del perfil en la fila de cabecera.
Private Function ResolverColumnasFacturas(ByRef per As TPerfil, ByRef datos As Variant, ByRef col() As Long, _
                                          ByRef msgError As String) As Boolean
    Dim i As Long, spec As String, opcional As Boolean, faltan As String
    If ColumnasMatriz(datos) = 0 Then
        msgError = "El origen no tiene datos."
        Exit Function
    End If
    If UBound(datos, 1) < per.FilaCabecera Then
        msgError = "El origen no tiene la fila de cabecera (" & per.FilaCabecera & ")."
        Exit Function
    End If
    For i = 1 To NUM_CAMPOS
        col(i) = 0
        spec = PerfilCampo(per, i)
        opcional = (Left$(spec, 1) = "?")
        If opcional Then spec = Mid$(spec, 2)
        If spec <> "" Then
            col(i) = BuscarColumna(datos, per.FilaCabecera, spec)
            If col(i) = 0 And (Not opcional Or CampoObligatorio(i)) Then
                faltan = faltan & vbCrLf & "   - " & CampoClave(i) & ": """ & Replace(spec, "|", """ o """) & """"
            End If
        ElseIf CampoObligatorio(i) Then
            faltan = faltan & vbCrLf & "   - " & CampoClave(i) & ": (sin configurar en el perfil)"
        End If
    Next i
    If faltan <> "" Then
        msgError = "No encuentro estas columnas del perfil """ & per.Nombre & """ en la fila " & per.FilaCabecera & _
                   " del origen:" & faltan & vbCrLf & vbCrLf & "Columnas del origen: " & ListaCabeceras(datos, per.FilaCabecera)
        Exit Function
    End If
    ResolverColumnasFacturas = True
End Function

Private Function ListaCabeceras(ByRef datos As Variant, ByVal fila As Long) As String
    Dim c As Long, s As String
    For c = LBound(datos, 2) To UBound(datos, 2)
        If ValorTexto(datos(fila, c)) <> "" Then
            If s <> "" Then s = s & ", "
            s = s & ValorTexto(datos(fila, c))
        End If
    Next c
    ListaCabeceras = s
End Function

Private Function Celda(ByRef datos As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c = 0 Then
        Celda = Empty
    Else
        Celda = datos(r, c)
    End If
End Function

' =====================================================================
'  LECTURA DE LAS LÍNEAS
' =====================================================================
Private Sub LeerLineasFacturas(ByRef emp As TEmpresa, ByRef per As TPerfil, ByRef datos As Variant, ByRef col() As Long)
    Dim r As Long, n As Long, ok As Boolean, errores As String, t As String
    Dim v As Variant, pctRECfg As Currency, ctaRECfg As String, ctaIVACfg As String, dummy As Currency
    Dim k As Long, avisos As String, pctCalc As Currency

    For r = per.FilaCabecera + 1 To UBound(datos, 1)
        If Not FilaVacia(datos, r) Then
            gNLinFac = gNLinFac + 1
            If gNLinFac > UBound(gLinFac) Then
                ReDim Preserve gLinFac(1 To UBound(gLinFac) * 2)
            End If
            n = gNLinFac
            errores = ""
            avisos = ""
            gLinFac(n).Fila = r
            For k = 1 To NUM_CAMPOS
                If col(k) > 0 Then
                    If IsError(datos(r, col(k))) Then errores = errores & "; error de fórmula (#N/D, #¡VALOR!...) en " & CampoClave(k)
                End If
            Next k
            If col(CAMPO_DOCUMENTO) > 0 Then
                If VarType(datos(r, col(CAMPO_DOCUMENTO))) = vbDate Then avisos = avisos & "; el nº de factura tiene formato de fecha en el origen"
            End If

            ' --- fecha y documento
            gLinFac(n).Fecha = LeerFecha(Celda(datos, r, col(CAMPO_FECHA)), per.FormatoFecha, ok)
            gLinFac(n).FechaOk = ok
            If ok Then
                gLinFac(n).ClaveFecha = FechaA3(gLinFac(n).Fecha)
            Else
                gLinFac(n).ClaveFecha = "?" & UCase$(ValorTexto(Celda(datos, r, col(CAMPO_FECHA))))
                errores = errores & "; fecha no válida [" & ValorTexto(Celda(datos, r, col(CAMPO_FECHA))) & "]"
            End If
            gLinFac(n).Documento = ValorTexto(Celda(datos, r, col(CAMPO_DOCUMENTO)))
            If gLinFac(n).Documento = "" Then errores = errores & "; número de factura vacío"
            gLinFac(n).FechaOp = gLinFac(n).Fecha
            If col(CAMPO_FECHA_OP) > 0 Then
                If Not EsVacio(datos(r, col(CAMPO_FECHA_OP))) Then
                    gLinFac(n).FechaOp = LeerFecha(datos(r, col(CAMPO_FECHA_OP)), per.FormatoFecha, ok)
                    If Not ok Then errores = errores & "; fecha de operación no válida"
                End If
            End If

            ' --- tipo y anulación
            gLinFac(n).TipoTexto = ValorTexto(Celda(datos, r, col(CAMPO_TIPO)))
            If EnLista(gLinFac(n).TipoTexto, per.ValoresAbono) Then
                gLinFac(n).Tipo = TIPO_ABONO
            ElseIf EnLista(gLinFac(n).TipoTexto, per.ValoresRectificativa) Then
                gLinFac(n).Tipo = TIPO_RECTIFICATIVA
            ElseIf EnLista(gLinFac(n).TipoTexto, per.ValoresTicket) Then
                gLinFac(n).Tipo = TIPO_TICKET
            Else
                gLinFac(n).Tipo = TIPO_FACTURA
            End If
            t = ValorTexto(Celda(datos, r, col(CAMPO_ANULADA)))
            gLinFac(n).Anulada = (t <> "" And Not EnLista(t, "NO;N;0;FALSO;FALSE"))

            ' --- cliente
            gLinFac(n).Cliente = ValorTexto(Celda(datos, r, col(CAMPO_CLIENTE)))
            gLinFac(n).NIF = ValorTexto(Celda(datos, r, col(CAMPO_NIF)))
            If gLinFac(n).NIF = "" Then gLinFac(n).NIF = ValorTexto(Celda(datos, r, col(CAMPO_NIF2)))
            v = Celda(datos, r, col(CAMPO_CP))
            gLinFac(n).CP = ValorTexto(v)
            If VarType(v) <> vbString And Len(gLinFac(n).CP) = 4 And SoloDigitos(gLinFac(n).CP) Then
                gLinFac(n).CP = "0" & gLinFac(n).CP               ' celda numérica: se había comido el cero
            End If

            ' --- importes
            gLinFac(n).BaseImp = LeerImporte(Celda(datos, r, col(CAMPO_BASE)), per.SepDecimal, ok)
            If Not ok Then errores = errores & "; base no numérica [" & ValorTexto(Celda(datos, r, col(CAMPO_BASE))) & "]"
            v = Celda(datos, r, col(CAMPO_PCT_IVA))
            gLinFac(n).PctIVA = LeerPorcentaje(v, per.SepDecimal, 1, ok)
            If Not ok Then errores = errores & "; % IVA no numérico [" & ValorTexto(v) & "]"
            gLinFac(n).Cuota = LeerImporte(Celda(datos, r, col(CAMPO_CUOTA_IVA)), per.SepDecimal, ok)
            If Not ok Then errores = errores & "; cuota no numérica [" & ValorTexto(Celda(datos, r, col(CAMPO_CUOTA_IVA))) & "]"
            gLinFac(n).PctRE = LeerPorcentaje(Celda(datos, r, col(CAMPO_PCT_RE)), per.SepDecimal, 0.1, ok)
            If Not ok Then errores = errores & "; % recargo no numérico"
            gLinFac(n).CuotaRE = LeerImporte(Celda(datos, r, col(CAMPO_CUOTA_RE)), per.SepDecimal, ok)
            If Not ok Then errores = errores & "; cuota de recargo no numérica"
            gLinFac(n).PctRet = LeerPorcentaje(Celda(datos, r, col(CAMPO_PCT_RET)), per.SepDecimal, 1, ok)
            If Not ok Then errores = errores & "; % retención no numérico"
            gLinFac(n).CuotaRet = LeerImporte(Celda(datos, r, col(CAMPO_CUOTA_RET)), per.SepDecimal, ok)
            If Not ok Then errores = errores & "; retención no numérica"
            ' retención exportada en negativo (total = base + IVA + retención): se pasa a positivo
            If gLinFac(n).CuotaRet <> 0 And gLinFac(n).BaseImp <> 0 Then
                If Sgn(gLinFac(n).CuotaRet) <> Sgn(gLinFac(n).BaseImp) Then
                    gLinFac(n).CuotaRet = -gLinFac(n).CuotaRet
                    avisos = avisos & "; retención con el signo cambiado en el origen: se toma como " & ImporteTexto(gLinFac(n).CuotaRet)
                End If
            End If
            ' retención sin porcentaje: se calcula si da un tipo habitual
            If gLinFac(n).CuotaRet <> 0 And gLinFac(n).PctRet = 0 And gLinFac(n).BaseImp <> 0 Then
                pctCalc = Redondear2(CDbl(gLinFac(n).CuotaRet) / CDbl(gLinFac(n).BaseImp) * 100#)
                If EnLista(PorcentajeTexto(pctCalc), "1;2;7;15;19;24") Then
                    gLinFac(n).PctRet = pctCalc
                    avisos = avisos & "; retención sin %: se calcula " & PorcentajeTexto(pctCalc) & "%"
                Else
                    errores = errores & "; retención sin % (el calculado, " & PorcentajeTexto(pctCalc) & "%, no es un tipo habitual)"
                End If
            End If
            v = Celda(datos, r, col(CAMPO_TOTAL))
            If EsVacio(v) Then
                gLinFac(n).Total = gLinFac(n).BaseImp + gLinFac(n).Cuota + gLinFac(n).CuotaRE - gLinFac(n).CuotaRet
            Else
                gLinFac(n).Total = LeerImporte(v, per.SepDecimal, ok)
                If Not ok Then errores = errores & "; total no numérico [" & ValorTexto(v) & "]"
            End If

            ' --- cuenta del cliente
            t = ValorTexto(Celda(datos, r, col(CAMPO_CTA_CLIENTE)))
            If t <> "" Then
                gLinFac(n).CtaCliente = t
                gLinFac(n).DescCtaCliente = ValorTexto(Celda(datos, r, col(CAMPO_DESC_CTA_CLIENTE)))
                If gLinFac(n).DescCtaCliente = "" Then gLinFac(n).DescCtaCliente = gLinFac(n).Cliente
                If gLinFac(n).DescCtaCliente = "" Then gLinFac(n).DescCtaCliente = "Cliente " & t
            Else
                gLinFac(n).CtaCliente = emp.CtaClientes
                gLinFac(n).DescCtaCliente = emp.DescClientes
            End If

            ' --- cuenta de ventas
            t = ValorTexto(Celda(datos, r, col(CAMPO_CTA_VENTAS)))
            If t <> "" Then
                gLinFac(n).CtaVentas = t
                gLinFac(n).DescCtaVentas = ValorTexto(Celda(datos, r, col(CAMPO_DESC_CTA_VENTAS)))
                If gLinFac(n).DescCtaVentas = "" Then
                    If t = emp.CtaVentasAlt And emp.CtaVentasAlt <> "" Then
                        gLinFac(n).DescCtaVentas = emp.DescVentasAlt
                    ElseIf t = emp.CtaVentas Then
                        gLinFac(n).DescCtaVentas = emp.DescVentas
                    Else
                        gLinFac(n).DescCtaVentas = "Ventas"
                    End If
                End If
            ElseIf col(CAMPO_SELECTOR) > 0 And emp.CtaVentasAlt <> "" And _
                   ValorTexto(Celda(datos, r, col(CAMPO_SELECTOR))) <> "" Then
                gLinFac(n).CtaVentas = emp.CtaVentasAlt
                gLinFac(n).DescCtaVentas = emp.DescVentasAlt
            Else
                gLinFac(n).CtaVentas = emp.CtaVentas
                gLinFac(n).DescCtaVentas = emp.DescVentas
            End If

            ' --- cuentas de IVA, recargo y retención
            ctaIVACfg = "": ctaRECfg = "": pctRECfg = 0
            IvaBuscar gLinFac(n).PctIVA, ctaIVACfg, pctRECfg, ctaRECfg
            t = ValorTexto(Celda(datos, r, col(CAMPO_CTA_IVA)))
            If t <> "" Then
                gLinFac(n).CtaIVA = t
            Else
                gLinFac(n).CtaIVA = ctaIVACfg
            End If
            gLinFac(n).CtaRE = ctaRECfg
            If gLinFac(n).CuotaRE <> 0 And gLinFac(n).PctRE = 0 Then gLinFac(n).PctRE = pctRECfg
            If gLinFac(n).CuotaRE <> 0 And gLinFac(n).PctRE = 0 Then errores = errores & "; recargo de equivalencia sin %"
            gLinFac(n).CtaRet = emp.CtaRetencion

            ' --- subtipo a3
            t = ValorTexto(Celda(datos, r, col(CAMPO_SUBTIPO)))
            If t = "" Then
                gLinFac(n).Subtipo = "01"
                If gLinFac(n).PctIVA = 0 And gLinFac(n).BaseImp <> 0 Then
                    avisos = avisos & "; IVA 0% sin subtipo: se envía 01 (interior sujeta). Indica 02, 03, 06, 08 o 09 si es exenta, " & _
                             "intracomunitaria, exportación o inversión del sujeto pasivo"
                End If
            ElseIf SoloDigitos(t) And Len(t) <= 2 Then
                gLinFac(n).Subtipo = Right$("0" & t, 2)
                If Not EnLista(gLinFac(n).Subtipo, "01;02;03;04;05;06;08;09") Then
                    errores = errores & "; subtipo no válido para ventas [" & t & "] (01-06, 08 o 09)"
                End If
            Else
                errores = errores & "; subtipo no válido [" & t & "]"
            End If
            t = ValorTexto(Celda(datos, r, col(CAMPO_IMPRESO)))
            If t = "" Then
                If gLinFac(n).Subtipo = "03" Or gLinFac(n).Subtipo = "04" Then
                    gLinFac(n).Impreso = "02"                   ' entregas intracomunitarias: 349 (bienes)
                Else
                    gLinFac(n).Impreso = "01"                   ' 347
                End If
            ElseIf SoloDigitos(t) And Len(t) <= 2 Then
                gLinFac(n).Impreso = Right$("0" & t, 2)
                If gLinFac(n).Impreso = "00" Then errores = errores & "; impreso no válido [" & t & "]"
            Else
                errores = errores & "; impreso no válido [" & t & "]"
            End If

            If errores <> "" Then gLinFac(n).ErrorLectura = Mid$(errores, 3)
            If avisos <> "" Then gLinFac(n).AvisoLectura = Mid$(avisos, 3)
            gLinFac(n).DentroFiltro = True
        End If
    Next r
End Sub

' =====================================================================
'  FILTROS
' =====================================================================
Private Sub AplicarFiltrosFacturas(ByRef fil As TFiltros, ByRef res As TResumen)
    Dim i As Long, dentro As Boolean
    res.FilasLeidas = gNLinFac
    For i = 1 To gNLinFac
        dentro = True
        With gLinFac(i)
            If .FechaOk Then
                If fil.UsarDesde Then If .Fecha < fil.FechaDesde Then dentro = False
                If fil.UsarHasta Then If .Fecha > fil.FechaHasta Then dentro = False
            End If
            If dentro Then dentro = DocumentoEnRango(.Documento, fil.RefDesde, fil.RefHasta)
            If dentro Then dentro = SerieAdmitida(.Documento, fil.SeriesIncluir, fil.SeriesExcluir)
            .DentroFiltro = dentro
            If Not dentro Then
                .Estado = "Fuera del filtro"
                res.FilasFiltradas = res.FilasFiltradas + 1
            Else
                res.TotalOrigen = res.TotalOrigen + .Total
            End If
        End With
    Next i
End Sub

' =====================================================================
'  AGRUPACIÓN POR DOCUMENTO (nº de factura + fecha, en orden de aparición)
' =====================================================================
Private Sub AgruparDocumentosFacturas()
    Dim i As Long, clave As String, claveBase As String, idx As Long, indice As New Collection, anuladas As New Collection
    Dim d As Long, v As Variant
    For i = 1 To gNLinFac
        If gLinFac(i).DentroFiltro Then
            claveBase = UCase$(gLinFac(i).Documento) & "|" & gLinFac(i).ClaveFecha
            ' una anulación y una nueva emisión con el mismo nº y fecha no se mezclan
            If gLinFac(i).Anulada Then
                clave = claveBase & "|ANULADA"
            Else
                clave = claveBase
            End If
            idx = 0
            On Error Resume Next
            idx = indice(clave)
            On Error GoTo 0
            If idx = 0 Then
                gNDocFac = gNDocFac + 1
                If gNDocFac > UBound(gDocFac) Then
                    ReDim Preserve gDocFac(1 To UBound(gDocFac) * 2)
                End If
                idx = gNDocFac
                indice.Add idx, clave
                gDocFac(idx).Documento = gLinFac(i).Documento
                gDocFac(idx).Fecha = gLinFac(i).Fecha
                gDocFac(idx).Primera = i
                gDocFac(idx).ClaveBase = claveBase
                gDocFac(idx).Anulado = gLinFac(i).Anulada
                If gLinFac(i).Anulada Then
                    On Error Resume Next
                    anuladas.Add idx, claveBase
                    On Error GoTo 0
                End If
            Else
                gLinFac(gDocFac(idx).Ultima).Sig = i
            End If
            gDocFac(idx).Ultima = i
            gDocFac(idx).NLineas = gDocFac(idx).NLineas + 1
            gLinFac(i).Doc = idx
            gLinFac(i).Sig = 0
        End If
    Next i
    For d = 1 To gNDocFac
        If Not gDocFac(d).Anulado Then
            If LeerColeccion(anuladas, gDocFac(d).ClaveBase, v) Then gDocFac(d).ConAnulado = True
        End If
    Next d
End Sub

' Tipo del documento: el de su primera línea; sin tipo y con total negativo,
' abono (si el perfil lo indica).
Private Function TipoDocumento(ByRef per As TPerfil, ByVal d As Long) As Integer
    Dim i As Long, sinTipo As Boolean, suma As Currency
    i = gDocFac(d).Primera
    TipoDocumento = gLinFac(i).Tipo
    If TipoDocumento <> TIPO_FACTURA Or Not per.AbonoSiNegativo Then Exit Function
    sinTipo = True
    Do While i > 0
        If gLinFac(i).TipoTexto <> "" Then sinTipo = False
        If gLinFac(i).BaseImp <> 0 Or gLinFac(i).Cuota <> 0 Or gLinFac(i).CuotaRE <> 0 Or gLinFac(i).CuotaRet <> 0 Then
            suma = suma + gLinFac(i).Total
        End If
        i = gLinFac(i).Sig
    Loop
    If sinTipo And suma < 0 Then TipoDocumento = TIPO_ABONO
End Function

Private Function TipoAdmitido(ByRef fil As TFiltros, ByVal tipo As Integer) As Boolean
    Select Case tipo
        Case TIPO_TICKET: TipoAdmitido = fil.InclTickets
        Case TIPO_ABONO, TIPO_RECTIFICATIVA: TipoAdmitido = fil.InclAbonos
        Case Else: TipoAdmitido = fil.InclFacturas
    End Select
End Function

' El documento queda fuera por el filtro de tipos
Private Sub FiltrarDocumento(ByVal d As Long, ByRef res As TResumen)
    Dim i As Long
    i = gDocFac(d).Primera
    Do While i > 0
        gLinFac(i).DentroFiltro = False
        gLinFac(i).Estado = "Fuera del filtro"
        res.FilasFiltradas = res.FilasFiltradas + 1
        res.TotalOrigen = res.TotalOrigen - gLinFac(i).Total
        i = gLinFac(i).Sig
    Loop
End Sub

' =====================================================================
'  PROCESO DE UN DOCUMENTO
' =====================================================================
Private Sub ProcesarDocumentoFactura(ByRef emp As TEmpresa, ByRef per As TPerfil, ByVal d As Long, ByRef res As TResumen)
    Dim i As Long, doc As String, fecha As Date, fechaOk As Boolean, sumAbs As Currency, sumTotal As Currency
    Dim sumOrigen As Currency, motivo As String, texto As String, esAbono As Boolean, sinTipo As Boolean
    Dim factor As Integer, signo As Integer, cliente As String, nif As String, cp As String, desc As String
    Dim det() As TDetFac, nDet As Long, k As Long, j As Long, encontrado As Boolean, orden() As Long, m As Long, t As Long
    Dim sumDet As Currency, malos As String, total As Currency, cuenta As String, nLin As Long
    Dim ctaCli As String, descCli As String, lineaIMU As String, primera As Long, esperada As Currency
    Dim ceroConTotal As Currency, registro As String, tipoDoc As Integer, esRect As Boolean, clave As String, previo As Variant

    primera = gDocFac(d).Primera
    doc = gDocFac(d).Documento
    fecha = gDocFac(d).Fecha
    fechaOk = gLinFac(primera).FechaOk
    res.UnidadesLeidas = res.UnidadesLeidas + 1

    ' Suma de las líneas del documento (origen)
    i = primera
    Do While i > 0
        sumOrigen = sumOrigen + gLinFac(i).Total
        i = gLinFac(i).Sig
    Loop

    ' 1) Anulado
    i = primera
    Do While i > 0
        If gLinFac(i).Anulada Then
            ExcluirDocumento d, "Documento anulado", "Excluido", sumOrigen, res
            Exit Sub
        End If
        i = gLinFac(i).Sig
    Loop

    ' 2) Errores de lectura (importes no numéricos, fecha no válida...)
    texto = ""
    i = primera
    Do While i > 0
        If gLinFac(i).ErrorLectura <> "" Then
            texto = texto & IIf(texto = "", "", " | ") & "fila " & gLinFac(i).Fila & ": " & gLinFac(i).ErrorLectura
        End If
        i = gLinFac(i).Sig
    Loop
    If texto <> "" Then
        ExcluirDocumento d, "Datos no válidos (" & texto & ")", "Excluido - revisar", sumOrigen, res
        Exit Sub
    End If

    ' 3) Documento con importe 0
    i = primera
    Do While i > 0
        sumAbs = sumAbs + Abs(gLinFac(i).Total)
        i = gLinFac(i).Sig
    Loop
    If sumAbs = 0 Then
        ExcluirDocumento d, "Importe 0,00", "Excluido", sumOrigen, res
        Exit Sub
    End If

    ' 4) Quitar líneas a cero (base, cuotas y retención a 0)
    i = primera
    Do While i > 0
        With gLinFac(i)
            If .BaseImp = 0 And .Cuota = 0 And .CuotaRE = 0 And .CuotaRet = 0 Then
                .Estado = "Línea a cero (omitida)"
                If .Total <> 0 Then ceroConTotal = ceroConTotal + .Total
            Else
                nLin = nLin + 1
                sumTotal = sumTotal + .Total
            End If
            i = .Sig
        End With
    Loop
    If nLin = 0 Then
        ExcluirDocumento d, "Sin líneas con importe", "Excluido", sumOrigen, res
        Exit Sub
    End If

    ' 5) Tipos de IVA y cuentas
    malos = ""
    i = primera
    Do While i > 0
        With gLinFac(i)
            If .Estado = "" Then
                texto = ""
                If .PctIVA < 0 Or .PctIVA >= 100 Then
                    texto = PorcentajeTexto(.PctIVA) & "% (no válido)"
                ElseIf .PctIVA = 0 And .Cuota <> 0 Then
                    texto = "0% con cuota " & ImporteTexto(.Cuota)
                ElseIf .CtaIVA = "" Then
                    texto = PorcentajeTexto(.PctIVA) & "%"
                End If
                If texto <> "" Then
                    If InStr(1, malos, texto, vbTextCompare) = 0 Then malos = malos & IIf(malos = "", "", ", ") & texto
                End If
            End If
            i = .Sig
        End With
    Loop
    If malos <> "" Then
        ExcluirDocumento d, "Tipo de IVA no configurado: " & malos & " (base " & ImporteTexto(SumaCampo(d, 1)) & _
            ", cuota " & ImporteTexto(SumaCampo(d, 2)) & ", total " & ImporteTexto(sumTotal) & ")", _
            "Excluido - contabilizar a mano", sumOrigen, res
        Exit Sub
    End If

    ' 6) Validar y normalizar cuentas con los dígitos del plan
    texto = ""
    i = primera
    Do While i > 0
        With gLinFac(i)
            If .Estado = "" Then
                cuenta = NormalizarCuenta(.CtaCliente, emp.Digitos, motivo)
                If cuenta = "" Then
                    AgregarMotivo texto, "cuenta de cliente: " & motivo
                Else
                    .CtaCliente = cuenta
                End If
                cuenta = NormalizarCuenta(.CtaVentas, emp.Digitos, motivo)
                If cuenta = "" Then
                    AgregarMotivo texto, "cuenta de ventas: " & motivo
                Else
                    .CtaVentas = cuenta
                End If
                cuenta = NormalizarCuenta(.CtaIVA, emp.Digitos, motivo)
                If cuenta = "" Then
                    AgregarMotivo texto, "cuenta de IVA: " & motivo
                Else
                    .CtaIVA = cuenta
                End If
                If .CuotaRE <> 0 Then
                    cuenta = NormalizarCuenta(.CtaRE, emp.Digitos, motivo)
                    If cuenta = "" Then
                        AgregarMotivo texto, "cuenta de recargo (hoja IVA): " & motivo
                    Else
                        .CtaRE = cuenta
                    End If
                Else
                    .CtaRE = ""
                End If
                If .CuotaRet <> 0 Then
                    cuenta = NormalizarCuenta(.CtaRet, emp.Digitos, motivo)
                    If cuenta = "" Then
                        AgregarMotivo texto, "cuenta de retenciones (hoja EMPRESAS): " & motivo
                    Else
                        .CtaRet = cuenta
                    End If
                Else
                    .CtaRet = ""
                End If
            End If
            i = .Sig
        End With
    Loop
    If texto <> "" Then
        ExcluirDocumento d, "Cuenta no válida: " & texto, "Excluido - revisar configuración", sumOrigen, res
        Exit Sub
    End If

    ' 7) Cuadre: el total debe coincidir con base + cuotas - retención
    sumDet = 0
    i = primera
    Do While i > 0
        With gLinFac(i)
            If .Estado = "" Then sumDet = sumDet + .BaseImp + .Cuota + .CuotaRE - .CuotaRet
            i = .Sig
        End With
    Loop
    If sumDet <> sumTotal Then
        ExcluirDocumento d, "Descuadre: total " & ImporteTexto(sumTotal) & " y base + cuotas " & ImporteTexto(sumDet) & _
            " (diferencia " & ImporteTexto(sumTotal - sumDet) & ")", "Excluido - revisar en origen", sumOrigen, res
        Exit Sub
    End If

    ' 7 bis) nº de factura repetido en a3 (a3 solo guarda 10 caracteres)
    clave = DocumentoA3(doc) & "|" & FechaA3(fecha)
    If LeerColeccion(mNumerosA3, clave, previo) Then
        ExcluirDocumento d, "Nº de factura repetido en a3: " & doc & " y " & CStr(previo) & " quedan los dos como " & _
            DocumentoA3(doc) & " (a3 solo admite 10 caracteres)", "Excluido - acortar el nº en el origen", sumOrigen, res
        Exit Sub
    End If

    ' 8) Tipo de documento y signo
    tipoDoc = TipoDocumento(per, d)
    esAbono = (tipoDoc = TIPO_ABONO)
    esRect = (tipoDoc = TIPO_RECTIFICATIVA)
    factor = 1
    If esRect Then
        ' rectificativa: tipo 2 con el signo del origen (negativa disminuye, positiva aumenta)
        factor = -1
        If sumTotal > 0 Then
            IncAgregar fechaOk, fecha, doc, gLinFac(primera).Fila, INC_AVISO, _
                "Rectificativa que aumenta la factura original (" & ImporteTexto(sumTotal) & ")", _
                "Incluida como tipo 2 con importes negativos - verificar"
        End If
    ElseIf esAbono Then
        If sumTotal < 0 Then
            factor = -1
        ElseIf sumTotal > 0 Then
            IncAgregar fechaOk, fecha, doc, gLinFac(primera).Fila, INC_AVISO, _
                "Abono con importes en positivo en el origen", "Incluido como abono (resta ventas) - verificar"
            res.AjusteSignoAbonos = res.AjusteSignoAbonos + 2 * sumTotal
        End If
    ElseIf sumTotal < 0 Then
        IncAgregar fechaOk, fecha, doc, gLinFac(primera).Fila, INC_AVISO, _
            "Factura con total negativo (" & ImporteTexto(sumTotal) & ")", "Incluida en negativo - verificar"
    End If
    If esAbono Or esRect Then signo = -1 Else signo = 1
    total = factor * sumTotal
    mNumerosA3.Add doc, clave

    ' 9) Avisos informativos
    If ceroConTotal <> 0 Then
        res.TotalExcluido = res.TotalExcluido + ceroConTotal
        IncAgregar fechaOk, fecha, doc, gLinFac(primera).Fila, INC_AVISO, _
            "Líneas con base y cuota a 0 pero con total " & ImporteTexto(ceroConTotal), "Líneas omitidas - revisar"
    End If
    If Len(AsciiA3(doc)) > 10 Then
        IncAgregar fechaOk, fecha, doc, gLinFac(primera).Fila, INC_AVISO, _
            "Número de factura de más de 10 caracteres: en a3 queda como " & DocumentoA3(doc), "Incluido - verificar"
    End If
    If gDocFac(d).ConAnulado Then
        IncAgregar fechaOk, fecha, doc, gLinFac(primera).Fila, INC_AVISO, _
            "Mismo nº y fecha que un documento anulado: se exporta la parte no anulada", "Incluido - verificar"
    End If
    i = primera
    Do While i > 0
        If gLinFac(i).Estado = "" And gLinFac(i).AvisoLectura <> "" Then
            IncAgregar fechaOk, fecha, doc, gLinFac(i).Fila, INC_AVISO, gLinFac(i).AvisoLectura, "Incluido - verificar"
        End If
        i = gLinFac(i).Sig
    Loop
    i = primera
    Do While i > 0
        With gLinFac(i)
            If .Estado = "" And .PctIVA > 0 Then
                esperada = Redondear2(CDbl(.BaseImp) * CDbl(.PctIVA) / 100#)
                If Abs(esperada - .Cuota) > 0.02 + Abs(.BaseImp) * 0.0005 Then
                    IncAgregar fechaOk, fecha, doc, .Fila, INC_AVISO, _
                        "La cuota " & ImporteTexto(.Cuota) & " no cuadra con base x " & PorcentajeTexto(.PctIVA) & _
                        "% = " & ImporteTexto(esperada), "Incluido tal cual - verificar"
                End If
            End If
            If .Estado = "" And .PctRE > 0 Then
                esperada = Redondear2(CDbl(.BaseImp) * CDbl(.PctRE) / 100#)
                If Abs(esperada - .CuotaRE) > 0.02 + Abs(.BaseImp) * 0.0005 Then
                    IncAgregar fechaOk, fecha, doc, .Fila, INC_AVISO, _
                        "La cuota de recargo " & ImporteTexto(.CuotaRE) & " no cuadra con base x " & PorcentajeTexto(.PctRE) & _
                        "% = " & ImporteTexto(esperada), "Incluido tal cual - verificar"
                End If
            End If
            If .Estado = "" And .PctRet > 0 Then
                esperada = Redondear2(CDbl(.BaseImp) * CDbl(.PctRet) / 100#)
                If Abs(esperada - .CuotaRet) > 0.02 + Abs(.BaseImp) * 0.0005 Then
                    IncAgregar fechaOk, fecha, doc, .Fila, INC_AVISO, _
                        "La retención " & ImporteTexto(.CuotaRet) & " no cuadra con base x " & PorcentajeTexto(.PctRet) & _
                        "% = " & ImporteTexto(esperada), "Incluido tal cual - verificar"
                End If
            End If
            i = .Sig
        End With
    Loop

    ' 10) Datos de cabecera
    cliente = AsciiA3(gLinFac(primera).Cliente)
    If cliente = "" Then cliente = per.NombreVacio
    nif = NormalizarNIF(gLinFac(primera).NIF)
    cp = ""
    If nif <> "" Then
        If Not NIFValido(nif) Then
            IncAgregar fechaOk, fecha, doc, gLinFac(primera).Fila, INC_AVISO, _
                "NIF " & nif & " no es un NIF español válido (a3 lo trata como español en el 347)", "Incluido - revisar 347/349 en a3"
        End If
        cp = AsciiA3(gLinFac(primera).CP)
        If cp <> "" And Not (Len(cp) = 5 And SoloDigitos(cp)) Then
            IncAgregar fechaOk, fecha, doc, gLinFac(primera).Fila, INC_AVISO, _
                "Código postal no español o incompleto (" & cp & "): se deja en blanco", _
                "Incluido - si es español, escríbelo con sus 5 cifras"
            cp = ""
        End If
    End If
    desc = doc & " " & cliente
    ctaCli = gLinFac(primera).CtaCliente
    descCli = gLinFac(primera).DescCtaCliente

    ' 11) Detalle agregado por cuenta de ventas + tipo de IVA
    ReDim det(1 To 8)
    nDet = 0
    i = primera
    Do While i > 0
        With gLinFac(i)
            If .Estado = "" Then
                encontrado = False
                For k = 1 To nDet
                    If det(k).CtaVentas = .CtaVentas And det(k).DescVentas = .DescCtaVentas And det(k).PctIVA = .PctIVA _
                       And det(k).CtaIVA = .CtaIVA And det(k).PctRE = .PctRE And det(k).CtaRE = .CtaRE _
                       And det(k).PctRet = .PctRet And det(k).CtaRet = .CtaRet And det(k).Subtipo = .Subtipo _
                       And det(k).Impreso = .Impreso Then
                        encontrado = True
                        Exit For
                    End If
                Next k
                If Not encontrado Then
                    nDet = nDet + 1
                    If nDet > UBound(det) Then
                        ReDim Preserve det(1 To nDet * 2)
                    End If
                    k = nDet
                    det(k).CtaVentas = .CtaVentas
                    det(k).DescVentas = .DescCtaVentas
                    det(k).PctIVA = .PctIVA
                    det(k).CtaIVA = .CtaIVA
                    det(k).PctRE = .PctRE
                    det(k).CtaRE = .CtaRE
                    det(k).PctRet = .PctRet
                    det(k).CtaRet = .CtaRet
                    det(k).Subtipo = .Subtipo
                    det(k).Impreso = .Impreso
                End If
                det(k).BaseImp = det(k).BaseImp + .BaseImp
                det(k).Cuota = det(k).Cuota + .Cuota
                det(k).CuotaRE = det(k).CuotaRE + .CuotaRE
                det(k).CuotaRet = det(k).CuotaRet + .CuotaRet
                .Estado = "Exportada"
            End If
            i = .Sig
        End With
    Loop
    ' orden: cuenta de ventas, descripción, % IVA, % recargo, % retención (ordenación de índices)
    ReDim orden(1 To nDet)
    For k = 1 To nDet
        orden(k) = k
    Next k
    For k = 2 To nDet
        t = orden(k)
        j = k - 1
        Do While j >= 1
            If Not DetalleMayor(det(orden(j)), det(t)) Then Exit Do
            orden(j + 1) = orden(j)
            j = j - 1
        Loop
        orden(j + 1) = t
    Next k

    ' 12) Registros
    registro = RegistroCabeceraFactura(emp.Codigo, fecha, esAbono Or esRect, ctaCli, descCli, "1", doc, desc, total, _
                                       nif, IIf(nif <> "", cliente, ""), cp, gLinFac(primera).FechaOp, fecha)
    DatAgregar registro
    For m = 1 To nDet
        k = orden(m)
        If m = nDet Then lineaIMU = "U" Else lineaIMU = "M"
        registro = RegistroDetalleIVA(emp.Codigo, fecha, det(k).CtaVentas, det(k).DescVentas, "C", doc, lineaIMU, desc, _
            det(k).Subtipo, factor * det(k).BaseImp, det(k).PctIVA, factor * det(k).Cuota, det(k).PctRE, _
            factor * det(k).CuotaRE, det(k).PctRet, factor * det(k).CuotaRet, det(k).Impreso, " ", " ", _
            det(k).CtaIVA, det(k).CtaRE, det(k).CtaRet, "", "")
        DatAgregar registro
        ' control (abonos en negativo)
        gNCtrlFac = gNCtrlFac + 1
        If gNCtrlFac > UBound(gCtrlFac) Then
            ReDim Preserve gCtrlFac(1 To UBound(gCtrlFac) * 2)
        End If
        With gCtrlFac(gNCtrlFac)
            .Fecha = fecha
            .Documento = doc
            If esAbono Then
                .TipoTexto = "Abono"
            ElseIf esRect Then
                .TipoTexto = "Rectificativa"
            ElseIf gLinFac(primera).Tipo = TIPO_TICKET Then
                .TipoTexto = "Ticket"
            Else
                .TipoTexto = "Factura"
            End If
            .Cliente = cliente
            .NIF = nif
            .CtaCliente = ctaCli
            .CtaVentas = det(k).CtaVentas
            .PctIVA = det(k).PctIVA
            .CtaIVA = det(k).CtaIVA
            .BaseImp = signo * factor * det(k).BaseImp
            .Cuota = signo * factor * det(k).Cuota
            .CuotaRE = signo * factor * det(k).CuotaRE
            .CuotaRet = signo * factor * det(k).CuotaRet
            .Total = .BaseImp + .Cuota + .CuotaRE - .CuotaRet
            res.BaseImp = res.BaseImp + .BaseImp
            res.Cuota = res.Cuota + .Cuota
            res.CuotaRE = res.CuotaRE + .CuotaRE
            res.CuotaRet = res.CuotaRet + .CuotaRet
            res.Total = res.Total + .Total
        End With
    Next m

    ' 13) Estadísticas
    res.UnidadesExportadas = res.UnidadesExportadas + 1
    If esAbono Or esRect Then
        res.NumAbonos = res.NumAbonos + 1
    ElseIf gLinFac(primera).Tipo = TIPO_TICKET Then
        res.NumTickets = res.NumTickets + 1
    Else
        res.NumFacturas = res.NumFacturas + 1
    End If
    If Not res.HayFechas Then
        res.FechaMin = fecha: res.FechaMax = fecha: res.HayFechas = True
    Else
        If fecha < res.FechaMin Then res.FechaMin = fecha
        If fecha > res.FechaMax Then res.FechaMax = fecha
    End If
End Sub

Private Function DetalleMayor(ByRef a As TDetFac, ByRef b As TDetFac) As Boolean
    If a.CtaVentas <> b.CtaVentas Then DetalleMayor = (a.CtaVentas > b.CtaVentas): Exit Function
    If a.DescVentas <> b.DescVentas Then DetalleMayor = (a.DescVentas > b.DescVentas): Exit Function
    If a.PctIVA <> b.PctIVA Then DetalleMayor = (a.PctIVA > b.PctIVA): Exit Function
    If a.PctRE <> b.PctRE Then DetalleMayor = (a.PctRE > b.PctRE): Exit Function
    If a.PctRet <> b.PctRet Then DetalleMayor = (a.PctRet > b.PctRet): Exit Function
    DetalleMayor = False
End Function

Private Sub AgregarMotivo(ByRef texto As String, ByVal motivo As String)
    If InStr(1, texto, motivo, vbTextCompare) = 0 Then
        If texto <> "" Then texto = texto & "; "
        texto = texto & motivo
    End If
End Sub

' Suma base (1) o cuota (2) de las líneas vivas del documento
Private Function SumaCampo(ByVal d As Long, ByVal campo As Integer) As Currency
    Dim i As Long, s As Currency
    i = gDocFac(d).Primera
    Do While i > 0
        If gLinFac(i).Estado = "" Then
            If campo = 1 Then s = s + gLinFac(i).BaseImp Else s = s + gLinFac(i).Cuota
        End If
        i = gLinFac(i).Sig
    Loop
    SumaCampo = s
End Function

Private Sub ExcluirDocumento(ByVal d As Long, ByVal texto As String, ByVal tratamiento As String, _
                             ByVal sumOrigen As Currency, ByRef res As TResumen)
    Dim i As Long, primera As Long
    primera = gDocFac(d).Primera
    IncAgregar gLinFac(primera).FechaOk, gDocFac(d).Fecha, gDocFac(d).Documento, gLinFac(primera).Fila, _
               INC_EXCLUIDO, texto, tratamiento
    i = primera
    Do While i > 0
        gLinFac(i).Estado = "Excluida"
        i = gLinFac(i).Sig
    Loop
    res.UnidadesExcluidas = res.UnidadesExcluidas + 1
    res.TotalExcluido = res.TotalExcluido + sumOrigen
End Sub
