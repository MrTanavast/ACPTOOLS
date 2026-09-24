Attribute VB_Name = "modA3Control"
' =====================================================================
'  ENLACE CONTABLE A3  ·  ACP ASOCIADOS
'  Módulo EXCEL DE CONTROL: crea un libro nuevo con el resumen, el
'  detalle de lo que va al SUENLACE.DAT, las incidencias, las filas de
'  origen con su estado y el propio fichero línea a línea.
' =====================================================================
Option Explicit

Private Const XL_CENTRO As Long = -4108
Private Const XL_IZQUIERDA As Long = -4131
Private Const XL_DERECHA As Long = -4152
Private Const XL_OPENXML As Long = 51

' Crea el libro de control. Si rutaGuardar no está vacía, lo guarda ahí.
Public Function CrearControl(ByRef p As TParametros, ByRef emp As TEmpresa, ByRef fil As TFiltros, _
                             ByRef res As TResumen, ByVal rutaDat As String, ByVal rutaGuardar As String) As Object
    Dim wb As Object, pantalla As Boolean, alertas As Boolean
    pantalla = Application.ScreenUpdating
    alertas = Application.DisplayAlerts
    Application.ScreenUpdating = False
    On Error GoTo Fallo

    Set wb = Application.Workbooks.Add
    ' dejar una sola hoja
    Application.DisplayAlerts = False
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = alertas

    HojaResumen wb.Worksheets(1), p, emp, fil, res, rutaDat
    If p.Modo = MODO_DIARIO Then
        HojaDetalleDiario NuevaHojaControl(wb, "Detalle SUENLACE")
    Else
        HojaDetalleFacturas NuevaHojaControl(wb, "Detalle SUENLACE")
    End If
    HojaIncidencias NuevaHojaControl(wb, "Incidencias")
    If p.Modo = MODO_DIARIO Then
        HojaOrigenDiario NuevaHojaControl(wb, "Origen")
    Else
        HojaOrigenFacturas NuevaHojaControl(wb, "Origen")
    End If
    HojaFichero NuevaHojaControl(wb, "Fichero DAT")
    wb.Worksheets(1).Activate
    wb.Worksheets(1).Range("A1").Select

    If rutaGuardar <> "" Then
        Application.DisplayAlerts = False
        wb.SaveAs Filename:=rutaGuardar, FileFormat:=XL_OPENXML
        Application.DisplayAlerts = alertas
    End If
    Application.ScreenUpdating = pantalla
    Set CrearControl = wb
    Exit Function
Fallo:
    Application.DisplayAlerts = alertas
    Application.ScreenUpdating = pantalla
    Err.Raise Err.Number, "CrearControl", Err.Description
End Function

Private Function NuevaHojaControl(ByVal wb As Object, ByVal nombre As String) As Object
    Dim ws As Object
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = nombre
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    Set NuevaHojaControl = ws
End Function

' Texto seguro para una celda (evita que "=..." o "-..." se tomen como fórmula)
Private Function TxtCelda(ByVal s As String) As String
    If Len(s) > 0 Then
        If InStr(1, "=+-@", Left$(s, 1)) > 0 Then s = "'" & s
    End If
    TxtCelda = s
End Function

Private Sub Cabecera(ByVal ws As Object, ByVal fila As Long, ByVal titulos As Variant)
    Dim i As Long
    For i = 0 To UBound(titulos)
        ws.Cells(fila, i + 1).Value = titulos(i)
    Next i
    EstiloCabecera ws.Range(ws.Cells(fila, 1), ws.Cells(fila, UBound(titulos) + 1))
    ws.Rows(fila).RowHeight = 30
End Sub

Private Sub Anchos(ByVal ws As Object, ByVal anchos As Variant)
    Dim i As Long
    For i = 0 To UBound(anchos)
        ws.Columns(i + 1).ColumnWidth = anchos(i)
    Next i
End Sub

Private Sub FilaTotales(ByVal ws As Object, ByVal filaDatos1 As Long, ByVal filaUltima As Long, ByVal columnas As Variant)
    Dim i As Long, f As Long, c As Long
    f = filaUltima + 2
    ws.Cells(f, 1).Value = "TOTAL (filtrado)"
    For i = 0 To UBound(columnas)
        c = columnas(i)
        If filaUltima >= filaDatos1 Then
            ws.Cells(f, c).Formula = "=SUBTOTAL(9," & ws.Range(ws.Cells(filaDatos1, c), ws.Cells(filaUltima, c)).Address(False, False) & ")"
        Else
            ws.Cells(f, c).Value = 0
        End If
        ws.Cells(f, c).NumberFormat = "#,##0.00"
    Next i
    With ws.Range(ws.Cells(f, 1), ws.Cells(f, columnas(UBound(columnas))))
        .Font.Bold = True
        .Interior.Color = COLOR_ACP_CLARO
        .Borders(8).Color = COLOR_ACP          ' xlEdgeTop
    End With
End Sub

' =====================================================================
'  HOJA RESUMEN
' =====================================================================
Private Sub HojaResumen(ByVal ws As Object, ByRef p As TParametros, ByRef emp As TEmpresa, ByRef fil As TFiltros, _
                        ByRef res As TResumen, ByVal rutaDat As String)
    Dim f As Long, dif As Currency
    ws.Name = "Resumen"
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    Anchos ws, Array(34, 18, 16, 16, 16, 16, 16)

    With ws.Range("A1:G1")
        .Merge
        .Value = "ACP ASOCIADOS  ·  Control del enlace contable a3"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = COLOR_BLANCO
        .Interior.Color = COLOR_ACP
        .HorizontalAlignment = XL_IZQUIERDA
        .IndentLevel = 1
    End With
    ws.Rows(1).RowHeight = 34
    With ws.Range("A2:G2")
        .Merge
        .Value = "Empresa " & emp.Codigo & IIf(emp.Nombre <> "", " - " & emp.Nombre, "") & "   ·   " & NombreModo(p)
        .Font.Size = 11
        .Font.Color = COLOR_BLANCO
        .Interior.Color = COLOR_ACP_OSCURO
        .IndentLevel = 1
    End With
    ws.Rows(2).RowHeight = 22

    f = 4
    f = Bloque(ws, f, "DATOS DE LA GENERACIÓN")
    f = Dato(ws, f, "Generado el", FechaTexto(Date) & "  " & Format$(Time, "hh:mm"))
    f = Dato(ws, f, "Usuario", Application.UserName)
    f = Dato(ws, f, "Origen de los datos", p.OrigenDescripcion)
    f = Dato(ws, f, "Fichero generado", rutaDat)
    f = Dato(ws, f, "Dígitos del plan de cuentas", CStr(emp.Digitos))
    f = Dato(ws, f, "Filtros", TextoFiltros(p, fil))
    If res.HayFechas Then f = Dato(ws, f, "Fechas exportadas", FechaTexto(res.FechaMin) & " a " & FechaTexto(res.FechaMax))
    f = f + 1

    f = Bloque(ws, f, "RESULTADO")
    If p.Modo = MODO_DIARIO Then
        f = Dato(ws, f, "Apuntes leídos del origen", res.FilasLeidas)
        f = Dato(ws, f, "Apuntes fuera del filtro", res.FilasFiltradas)
        f = Dato(ws, f, "Asientos dentro del filtro", res.UnidadesLeidas)
        f = DatoColor(ws, f, "Asientos exportados", res.UnidadesExportadas, COLOR_OK)
        f = DatoColor(ws, f, "Asientos excluidos", res.UnidadesExcluidas, IIf(res.UnidadesExcluidas > 0, COLOR_ERROR, COLOR_TEXTO))
    Else
        f = Dato(ws, f, "Líneas leídas del origen", res.FilasLeidas)
        f = Dato(ws, f, "Líneas fuera del filtro", res.FilasFiltradas)
        f = Dato(ws, f, "Documentos dentro del filtro", res.UnidadesLeidas)
        f = DatoColor(ws, f, "Documentos exportados", res.UnidadesExportadas, COLOR_OK)
        f = Dato(ws, f, "   · Facturas", res.NumFacturas)
        f = Dato(ws, f, "   · Tickets", res.NumTickets)
        f = Dato(ws, f, "   · Abonos", res.NumAbonos)
        f = DatoColor(ws, f, "Documentos excluidos", res.UnidadesExcluidas, IIf(res.UnidadesExcluidas > 0, COLOR_ERROR, COLOR_TEXTO))
    End If
    f = Dato(ws, f, "Líneas del SUENLACE.DAT", res.LineasDat)
    f = DatoColor(ws, f, "Avisos", res.Avisos, IIf(res.Avisos > 0, COLOR_AVISO, COLOR_TEXTO))
    f = f + 1

    If p.Modo = MODO_DIARIO Then
        f = Bloque(ws, f, "IMPORTES EXPORTADOS")
        f = DatoImporte(ws, f, "Total Debe", res.Debe)
        f = DatoImporte(ws, f, "Total Haber", res.Haber)
        dif = res.Debe - res.Haber
        f = DatoImporte(ws, f, "Diferencia Debe - Haber", dif)
        ws.Cells(f - 1, 2).Font.Color = IIf(dif = 0, COLOR_OK, COLOR_ERROR)
        ws.Cells(f - 1, 2).Font.Bold = True
        f = f + 1
        f = ResumenPorCuentaDiario(ws, f)
    Else
        f = Bloque(ws, f, "IMPORTES EXPORTADOS (abonos en negativo)")
        f = DatoImporte(ws, f, "Base imponible", res.BaseImp)
        f = DatoImporte(ws, f, "Cuota de IVA", res.Cuota)
        If res.CuotaRE <> 0 Then f = DatoImporte(ws, f, "Recargo de equivalencia", res.CuotaRE)
        If res.CuotaRet <> 0 Then f = DatoImporte(ws, f, "Retenciones", res.CuotaRet)
        f = DatoImporte(ws, f, "Total", res.Total)
        f = f + 1
        f = Bloque(ws, f, "CUADRE CON EL ORIGEN")
        f = DatoImporte(ws, f, "Total del origen (dentro del filtro)", res.TotalOrigen)
        f = DatoImporte(ws, f, "- Excluido (anulados, incidencias...)", res.TotalExcluido)
        f = DatoImporte(ws, f, "- Abonos que venían en positivo (x2)", res.AjusteSignoAbonos)
        f = DatoImporte(ws, f, "- Exportado al SUENLACE", res.Total)
        dif = res.TotalOrigen - res.TotalExcluido - res.AjusteSignoAbonos - res.Total
        f = DatoImporte(ws, f, "= Diferencia", dif)
        ws.Cells(f - 1, 2).Font.Color = IIf(dif = 0, COLOR_OK, COLOR_ERROR)
        ws.Cells(f - 1, 2).Font.Bold = True
        ws.Cells(f - 1, 3).Value = IIf(dif = 0, "Cuadra", "NO CUADRA - revisar")
        ws.Cells(f - 1, 3).Font.Color = IIf(dif = 0, COLOR_OK, COLOR_ERROR)
        ws.Cells(f - 1, 3).Font.Bold = True
        f = f + 1
        f = ResumenPorCuentaFacturas(ws, f)
    End If
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Private Function NombreModo(ByRef p As TParametros) As String
    If p.Modo = MODO_DIARIO Then
        NombreModo = "Diario (asientos sin IVA)"
    Else
        NombreModo = "Libro de facturas emitidas (perfil " & p.Perfil & ")"
    End If
End Function

Public Function TextoFiltros(ByRef p As TParametros, ByRef fil As TFiltros) As String
    Dim s As String, tipos As String
    If fil.UsarDesde Then s = s & "desde " & FechaTexto(fil.FechaDesde) & "  "
    If fil.UsarHasta Then s = s & "hasta " & FechaTexto(fil.FechaHasta) & "  "
    If fil.RefDesde <> "" Or fil.RefHasta <> "" Then
        s = s & IIf(p.Modo = MODO_DIARIO, "asientos ", "documentos ") & _
            IIf(fil.RefDesde <> "", fil.RefDesde, "...") & " a " & IIf(fil.RefHasta <> "", fil.RefHasta, "...") & "  "
    End If
    If p.Modo <> MODO_DIARIO Then
        If fil.SeriesIncluir <> "" Then s = s & "series: " & fil.SeriesIncluir & "  "
        If fil.SeriesExcluir <> "" Then s = s & "sin series: " & fil.SeriesExcluir & "  "
        If Not (fil.InclFacturas And fil.InclTickets And fil.InclAbonos) Then
            If fil.InclFacturas Then tipos = tipos & "facturas "
            If fil.InclTickets Then tipos = tipos & "tickets "
            If fil.InclAbonos Then tipos = tipos & "abonos "
            s = s & "solo " & Trim$(tipos) & "  "
        End If
    ElseIf Not fil.ExcluirDescuadrados Then
        s = s & "incluye asientos descuadrados  "
    End If
    If s = "" Then s = "ninguno (todo el origen)"
    TextoFiltros = Trim$(s)
End Function

Private Function Bloque(ByVal ws As Object, ByVal f As Long, ByVal titulo As String) As Long
    With ws.Range(ws.Cells(f, 1), ws.Cells(f, 7))
        .Interior.Color = COLOR_ACP_CLARO
        .Font.Bold = True
        .Font.Color = COLOR_ACP
        .Borders(9).Color = COLOR_ACP        ' xlEdgeBottom
    End With
    ws.Cells(f, 1).Value = titulo
    Bloque = f + 1
End Function

Private Function Dato(ByVal ws As Object, ByVal f As Long, ByVal etiqueta As String, ByVal valor As Variant) As Long
    ws.Cells(f, 1).Value = etiqueta
    If VarType(valor) = vbString Then ws.Cells(f, 2).Value = TxtCelda(CStr(valor)) Else ws.Cells(f, 2).Value = valor
    ws.Cells(f, 2).HorizontalAlignment = XL_IZQUIERDA
    Dato = f + 1
End Function

Private Function DatoColor(ByVal ws As Object, ByVal f As Long, ByVal etiqueta As String, ByVal valor As Variant, ByVal color As Long) As Long
    DatoColor = Dato(ws, f, etiqueta, valor)
    ws.Cells(f, 2).Font.Color = color
    ws.Cells(f, 2).Font.Bold = True
End Function

Private Function DatoImporte(ByVal ws As Object, ByVal f As Long, ByVal etiqueta As String, ByVal valor As Currency) As Long
    ws.Cells(f, 1).Value = etiqueta
    ws.Cells(f, 2).Value = CDbl(valor)
    ws.Cells(f, 2).NumberFormat = "#,##0.00"
    ws.Cells(f, 2).HorizontalAlignment = XL_DERECHA
    DatoImporte = f + 1
End Function

' Resumen por cuenta de ventas y tipo de IVA (facturas)
Private Function ResumenPorCuentaFacturas(ByVal ws As Object, ByVal f As Long) As Long
    Dim claves() As String, ctas() As String, pcts() As Currency, ivas() As String
    Dim b() As Currency, c() As Currency, t() As Currency, n As Long, i As Long, k As Long, clave As String
    Dim orden() As Long, j As Long, x As Long, ini As Long
    f = Bloque(ws, f, "RESUMEN POR CUENTA DE VENTAS Y TIPO DE IVA")
    Cabecera ws, f, Array("Cuenta de ventas", "% IVA", "Cuenta IVA", "Base", "Cuota IVA", "Total")
    f = f + 1
    If gNCtrlFac = 0 Then ResumenPorCuentaFacturas = f + 1: Exit Function
    ReDim claves(1 To gNCtrlFac): ReDim ctas(1 To gNCtrlFac): ReDim pcts(1 To gNCtrlFac): ReDim ivas(1 To gNCtrlFac)
    ReDim b(1 To gNCtrlFac): ReDim c(1 To gNCtrlFac): ReDim t(1 To gNCtrlFac)
    For i = 1 To gNCtrlFac
        clave = gCtrlFac(i).CtaVentas & "|" & PorcentajeA3(gCtrlFac(i).PctIVA) & "|" & gCtrlFac(i).CtaIVA
        k = 0
        For j = 1 To n
            If claves(j) = clave Then k = j: Exit For
        Next j
        If k = 0 Then
            n = n + 1: k = n
            claves(k) = clave: ctas(k) = gCtrlFac(i).CtaVentas: pcts(k) = gCtrlFac(i).PctIVA: ivas(k) = gCtrlFac(i).CtaIVA
        End If
        b(k) = b(k) + gCtrlFac(i).BaseImp
        c(k) = c(k) + gCtrlFac(i).Cuota
        t(k) = t(k) + gCtrlFac(i).Total
    Next i
    ReDim orden(1 To n)
    For i = 1 To n: orden(i) = i: Next i
    For i = 2 To n
        x = orden(i): j = i - 1
        Do While j >= 1
            If claves(orden(j)) <= claves(x) Then Exit Do
            orden(j + 1) = orden(j): j = j - 1
        Loop
        orden(j + 1) = x
    Next i
    ini = f
    For i = 1 To n
        k = orden(i)
        ws.Cells(f, 1).Value = TxtCelda(ctas(k))
        ws.Cells(f, 2).Value = CDbl(pcts(k))
        ws.Cells(f, 3).Value = TxtCelda(ivas(k))
        ws.Cells(f, 4).Value = CDbl(b(k))
        ws.Cells(f, 5).Value = CDbl(c(k))
        ws.Cells(f, 6).Value = CDbl(t(k))
        f = f + 1
    Next i
    ws.Range(ws.Cells(ini, 4), ws.Cells(f, 6)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(ini, 2), ws.Cells(f, 2)).NumberFormat = "0.00"
    ws.Cells(f, 1).Value = "TOTAL"
    ws.Cells(f, 4).Formula = "=SUM(" & ws.Range(ws.Cells(ini, 4), ws.Cells(f - 1, 4)).Address(False, False) & ")"
    ws.Cells(f, 5).Formula = "=SUM(" & ws.Range(ws.Cells(ini, 5), ws.Cells(f - 1, 5)).Address(False, False) & ")"
    ws.Cells(f, 6).Formula = "=SUM(" & ws.Range(ws.Cells(ini, 6), ws.Cells(f - 1, 6)).Address(False, False) & ")"
    With ws.Range(ws.Cells(f, 1), ws.Cells(f, 6))
        .Font.Bold = True
        .Interior.Color = COLOR_ACP_CLARO
    End With
    ResumenPorCuentaFacturas = f + 2
End Function

' Resumen por cuenta (diario)
Private Function ResumenPorCuentaDiario(ByVal ws As Object, ByVal f As Long) As Long
    Dim ctas() As String, noms() As String, d() As Currency, h() As Currency, n As Long, i As Long, j As Long, k As Long
    Dim orden() As Long, x As Long, ini As Long
    f = Bloque(ws, f, "RESUMEN POR CUENTA")
    Cabecera ws, f, Array("Cuenta", "Nombre", "Debe", "Haber", "Saldo")
    f = f + 1
    If gNCtrlDia = 0 Then ResumenPorCuentaDiario = f + 1: Exit Function
    ReDim ctas(1 To gNCtrlDia): ReDim noms(1 To gNCtrlDia): ReDim d(1 To gNCtrlDia): ReDim h(1 To gNCtrlDia)
    For i = 1 To gNCtrlDia
        k = 0
        For j = 1 To n
            If ctas(j) = gCtrlDia(i).Cuenta Then k = j: Exit For
        Next j
        If k = 0 Then
            n = n + 1: k = n
            ctas(k) = gCtrlDia(i).Cuenta: noms(k) = gCtrlDia(i).NombreCuenta
        End If
        d(k) = d(k) + gCtrlDia(i).Debe
        h(k) = h(k) + gCtrlDia(i).Haber
    Next i
    ReDim orden(1 To n)
    For i = 1 To n: orden(i) = i: Next i
    For i = 2 To n
        x = orden(i): j = i - 1
        Do While j >= 1
            If ctas(orden(j)) <= ctas(x) Then Exit Do
            orden(j + 1) = orden(j): j = j - 1
        Loop
        orden(j + 1) = x
    Next i
    ini = f
    For i = 1 To n
        k = orden(i)
        ws.Cells(f, 1).Value = TxtCelda(ctas(k))
        ws.Cells(f, 2).Value = TxtCelda(noms(k))
        ws.Cells(f, 3).Value = CDbl(d(k))
        ws.Cells(f, 4).Value = CDbl(h(k))
        ws.Cells(f, 5).Value = CDbl(d(k) - h(k))
        f = f + 1
    Next i
    ws.Range(ws.Cells(ini, 3), ws.Cells(f, 5)).NumberFormat = "#,##0.00"
    ws.Cells(f, 1).Value = "TOTAL"
    ws.Cells(f, 3).Formula = "=SUM(" & ws.Range(ws.Cells(ini, 3), ws.Cells(f - 1, 3)).Address(False, False) & ")"
    ws.Cells(f, 4).Formula = "=SUM(" & ws.Range(ws.Cells(ini, 4), ws.Cells(f - 1, 4)).Address(False, False) & ")"
    ws.Cells(f, 5).Formula = "=SUM(" & ws.Range(ws.Cells(ini, 5), ws.Cells(f - 1, 5)).Address(False, False) & ")"
    With ws.Range(ws.Cells(f, 1), ws.Cells(f, 5))
        .Font.Bold = True
        .Interior.Color = COLOR_ACP_CLARO
    End With
    ResumenPorCuentaDiario = f + 2
End Function

' =====================================================================
'  HOJAS DE DETALLE
' =====================================================================
Private Sub HojaDetalleFacturas(ByVal ws As Object)
    Dim m() As Variant, i As Long
    Cabecera ws, 1, Array("Fecha", "Documento", "Tipo", "Cliente", "NIF", "Cta. cliente", "Cta. ventas", _
                          "% IVA", "Cta. IVA", "Base", "Cuota IVA", "Recargo", "Retención", "Total")
    Anchos ws, Array(11, 12, 9, 34, 12, 12, 12, 7, 12, 13, 12, 10, 10, 13)
    If gNCtrlFac > 0 Then
        ReDim m(1 To gNCtrlFac, 1 To 14)
        For i = 1 To gNCtrlFac
            m(i, 1) = gCtrlFac(i).Fecha
            m(i, 2) = TxtCelda(gCtrlFac(i).Documento)
            m(i, 3) = gCtrlFac(i).TipoTexto
            m(i, 4) = TxtCelda(gCtrlFac(i).Cliente)
            m(i, 5) = TxtCelda(gCtrlFac(i).NIF)
            m(i, 6) = TxtCelda(gCtrlFac(i).CtaCliente)
            m(i, 7) = TxtCelda(gCtrlFac(i).CtaVentas)
            m(i, 8) = CDbl(gCtrlFac(i).PctIVA)
            m(i, 9) = TxtCelda(gCtrlFac(i).CtaIVA)
            m(i, 10) = CDbl(gCtrlFac(i).BaseImp)
            m(i, 11) = CDbl(gCtrlFac(i).Cuota)
            m(i, 12) = CDbl(gCtrlFac(i).CuotaRE)
            m(i, 13) = CDbl(gCtrlFac(i).CuotaRet)
            m(i, 14) = CDbl(gCtrlFac(i).Total)
        Next i
        ws.Range(ws.Cells(2, 2), ws.Cells(gNCtrlFac + 1, 7)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 9), ws.Cells(gNCtrlFac + 1, 9)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 1), ws.Cells(gNCtrlFac + 1, 14)).Value = m
        ws.Range(ws.Cells(2, 8), ws.Cells(gNCtrlFac + 1, 8)).NumberFormat = "0.00"
        ws.Range(ws.Cells(2, 10), ws.Cells(gNCtrlFac + 1, 14)).NumberFormat = "#,##0.00"
    End If
    ws.Columns(1).NumberFormat = "dd/mm/yyyy"
    FilaTotales ws, 2, gNCtrlFac + 1, Array(10, 11, 12, 13, 14)
    TerminarTabla ws, 14, gNCtrlFac + 1
End Sub

Private Sub HojaDetalleDiario(ByVal ws As Object)
    Dim m() As Variant, i As Long
    Cabecera ws, 1, Array("Fecha", "Asiento", "Línea", "Cuenta", "Nombre cuenta", "Descripción", "Documento", "Debe", "Haber")
    Anchos ws, Array(11, 9, 6, 13, 32, 34, 12, 13, 13)
    If gNCtrlDia > 0 Then
        ReDim m(1 To gNCtrlDia, 1 To 9)
        For i = 1 To gNCtrlDia
            m(i, 1) = gCtrlDia(i).Fecha
            m(i, 2) = TxtCelda(gCtrlDia(i).Asiento)
            m(i, 3) = gCtrlDia(i).LineaApunte
            m(i, 4) = TxtCelda(gCtrlDia(i).Cuenta)
            m(i, 5) = TxtCelda(gCtrlDia(i).NombreCuenta)
            m(i, 6) = TxtCelda(gCtrlDia(i).Descripcion)
            m(i, 7) = TxtCelda(gCtrlDia(i).Documento)
            m(i, 8) = CDbl(gCtrlDia(i).Debe)
            m(i, 9) = CDbl(gCtrlDia(i).Haber)
        Next i
        ws.Range(ws.Cells(2, 4), ws.Cells(gNCtrlDia + 1, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 1), ws.Cells(gNCtrlDia + 1, 9)).Value = m
        ws.Range(ws.Cells(2, 8), ws.Cells(gNCtrlDia + 1, 9)).NumberFormat = "#,##0.00"
    End If
    ws.Columns(1).NumberFormat = "dd/mm/yyyy"
    FilaTotales ws, 2, gNCtrlDia + 1, Array(8, 9)
    TerminarTabla ws, 9, gNCtrlDia + 1
End Sub

Private Sub HojaIncidencias(ByVal ws As Object)
    Dim m() As Variant, i As Long, color As Long
    Cabecera ws, 1, Array("Gravedad", "Fecha", "Documento / asiento", "Fila origen", "Incidencia", "Tratamiento")
    Anchos ws, Array(11, 11, 18, 10, 90, 40)
    If gNInc > 0 Then
        ReDim m(1 To gNInc, 1 To 6)
        For i = 1 To gNInc
            m(i, 1) = gInc(i).Gravedad
            If gInc(i).TieneFecha Then m(i, 2) = gInc(i).Fecha Else m(i, 2) = Empty
            m(i, 3) = TxtCelda(gInc(i).Referencia)
            If gInc(i).Fila > 0 Then m(i, 4) = gInc(i).Fila Else m(i, 4) = Empty
            m(i, 5) = TxtCelda(gInc(i).Texto)
            m(i, 6) = TxtCelda(gInc(i).Tratamiento)
        Next i
        ws.Range(ws.Cells(2, 3), ws.Cells(gNInc + 1, 3)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 1), ws.Cells(gNInc + 1, 6)).Value = m
        For i = 1 To gNInc
            Select Case gInc(i).Gravedad
                Case INC_EXCLUIDO: color = COLOR_ERROR_CLARO
                Case INC_AVISO: color = COLOR_AVISO_CLARO
                Case Else: color = COLOR_ACP_CLARO
            End Select
            ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 6)).Interior.Color = color
        Next i
        ws.Range(ws.Cells(2, 5), ws.Cells(gNInc + 1, 5)).WrapText = True
    Else
        ws.Cells(2, 1).Value = "Sin incidencias"
        ws.Cells(2, 1).Font.Color = COLOR_OK
        ws.Cells(2, 1).Font.Bold = True
    End If
    ws.Columns(2).NumberFormat = "dd/mm/yyyy"
    TerminarTabla ws, 6, gNInc + 1
End Sub

Private Sub HojaOrigenFacturas(ByVal ws As Object)
    Dim m() As Variant, i As Long
    Cabecera ws, 1, Array("Fila", "Fecha", "Documento", "Tipo", "Cliente", "NIF", "Base", "% IVA", "Cuota IVA", _
                          "Total", "Cta. ventas", "Cta. IVA", "Estado", "Error de lectura")
    Anchos ws, Array(7, 11, 12, 12, 30, 12, 12, 7, 11, 12, 12, 12, 22, 40)
    If gNLinFac > 0 Then
        ReDim m(1 To gNLinFac, 1 To 14)
        For i = 1 To gNLinFac
            m(i, 1) = gLinFac(i).Fila
            If gLinFac(i).FechaOk Then m(i, 2) = gLinFac(i).Fecha Else m(i, 2) = Empty
            m(i, 3) = TxtCelda(gLinFac(i).Documento)
            m(i, 4) = TxtCelda(gLinFac(i).TipoTexto)
            m(i, 5) = TxtCelda(gLinFac(i).Cliente)
            m(i, 6) = TxtCelda(gLinFac(i).NIF)
            m(i, 7) = CDbl(gLinFac(i).BaseImp)
            m(i, 8) = CDbl(gLinFac(i).PctIVA)
            m(i, 9) = CDbl(gLinFac(i).Cuota)
            m(i, 10) = CDbl(gLinFac(i).Total)
            m(i, 11) = TxtCelda(gLinFac(i).CtaVentas)
            m(i, 12) = TxtCelda(gLinFac(i).CtaIVA)
            m(i, 13) = IIf(gLinFac(i).Estado = "", "Sin procesar", gLinFac(i).Estado)
            m(i, 14) = TxtCelda(gLinFac(i).ErrorLectura)
        Next i
        ws.Range(ws.Cells(2, 3), ws.Cells(gNLinFac + 1, 6)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 11), ws.Cells(gNLinFac + 1, 12)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 1), ws.Cells(gNLinFac + 1, 14)).Value = m
        ws.Range(ws.Cells(2, 7), ws.Cells(gNLinFac + 1, 10)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(2, 8), ws.Cells(gNLinFac + 1, 8)).NumberFormat = "0.00"
        ColorearEstados ws, 13, gNLinFac + 1
    End If
    ws.Columns(2).NumberFormat = "dd/mm/yyyy"
    FilaTotales ws, 2, gNLinFac + 1, Array(7, 9, 10)
    TerminarTabla ws, 14, gNLinFac + 1
End Sub

Private Sub HojaOrigenDiario(ByVal ws As Object)
    Dim m() As Variant, i As Long
    Cabecera ws, 1, Array("Fila", "Fecha", "Asiento", "Cuenta", "Nombre cuenta", "Descripción", "Documento", _
                          "D/H", "Importe", "Estado", "Error de lectura")
    Anchos ws, Array(7, 11, 9, 13, 30, 30, 12, 5, 13, 22, 50)
    If gNLinDia > 0 Then
        ReDim m(1 To gNLinDia, 1 To 11)
        For i = 1 To gNLinDia
            m(i, 1) = gLinDia(i).Fila
            If gLinDia(i).FechaOk Then m(i, 2) = gLinDia(i).Fecha Else m(i, 2) = Empty
            m(i, 3) = TxtCelda(gLinDia(i).Asiento)
            m(i, 4) = TxtCelda(gLinDia(i).Cuenta)
            m(i, 5) = TxtCelda(gLinDia(i).NombreCuenta)
            m(i, 6) = TxtCelda(gLinDia(i).Descripcion)
            m(i, 7) = TxtCelda(gLinDia(i).Documento)
            m(i, 8) = gLinDia(i).DH
            m(i, 9) = CDbl(gLinDia(i).Importe)
            m(i, 10) = IIf(gLinDia(i).Estado = "", "Sin procesar", gLinDia(i).Estado)
            m(i, 11) = TxtCelda(gLinDia(i).ErrorTxt)
        Next i
        ws.Range(ws.Cells(2, 3), ws.Cells(gNLinDia + 1, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 1), ws.Cells(gNLinDia + 1, 11)).Value = m
        ws.Range(ws.Cells(2, 9), ws.Cells(gNLinDia + 1, 9)).NumberFormat = "#,##0.00"
        ColorearEstados ws, 10, gNLinDia + 1
    End If
    ws.Columns(2).NumberFormat = "dd/mm/yyyy"
    TerminarTabla ws, 11, gNLinDia + 1
End Sub

Private Sub ColorearEstados(ByVal ws As Object, ByVal col As Long, ByVal ultima As Long)
    Dim i As Long, v As String
    For i = 2 To ultima
        v = CStr(ws.Cells(i, col).Value)
        If v = "Exportada" Then
            ws.Cells(i, col).Font.Color = COLOR_OK
        ElseIf v = "Excluida" Then
            ws.Cells(i, col).Font.Color = COLOR_ERROR
            ws.Cells(i, col).Font.Bold = True
        ElseIf v <> "" Then
            ws.Cells(i, col).Font.Color = COLOR_GRIS
        End If
    Next i
End Sub

Private Sub HojaFichero(ByVal ws As Object)
    Dim m() As Variant, i As Long
    Cabecera ws, 1, Array("Línea", "Registro", "Documento", "Contenido (254 posiciones)")
    Anchos ws, Array(7, 9, 12, 200)
    If gNDat > 0 Then
        ReDim m(1 To gNDat, 1 To 4)
        For i = 1 To gNDat
            m(i, 1) = i
            m(i, 2) = Mid$(gDat(i), 15, 1)
            m(i, 3) = TxtCelda(Trim$(Mid$(gDat(i), 59, 10)))
            m(i, 4) = "'" & gDat(i)
        Next i
        ws.Range(ws.Cells(2, 2), ws.Cells(gNDat + 1, 3)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 1), ws.Cells(gNDat + 1, 4)).Value = m
        ws.Range(ws.Cells(2, 4), ws.Cells(gNDat + 1, 4)).Font.Name = "Consolas"
        ws.Range(ws.Cells(2, 4), ws.Cells(gNDat + 1, 4)).Font.Size = 9
    End If
    ws.Cells(gNDat + 3, 4).Value = "Regla: 1 tipo formato | 2-6 empresa | 7-14 fecha | 15 registro | 16-27 cuenta | 28-57 descripción | 58 D/H o tipo | 59-68 documento | 69 I/M/U | 70-99 concepto | 100+ importes"
    ws.Cells(gNDat + 3, 4).Font.Color = COLOR_GRIS
    TerminarTabla ws, 4, gNDat + 1
End Sub

Private Sub TerminarTabla(ByVal ws As Object, ByVal nCols As Long, ByVal ultima As Long)
    If ultima < 2 Then ultima = 2
    ws.Range(ws.Cells(1, 1), ws.Cells(ultima, nCols)).AutoFilter
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Tab.Color = COLOR_ACP
End Sub
