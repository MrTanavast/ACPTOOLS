Attribute VB_Name = "modA3Inicio"
' =====================================================================
'  ENLACE CONTABLE A3  ·  ACP ASOCIADOS
'  Módulo INICIO: macros que se lanzan (Alt+F8 o botones de la hoja
'  INICIO) y la orquestación de cada generación.
'
'    EnlaceA3               -> abre la ventana principal
'    PrepararLibro          -> crea las hojas INICIO, EMPRESAS, IVA,
'                              PERFILES y las plantillas (una sola vez)
'    NuevaPlantillaDiario   -> libro nuevo con la plantilla de diario
'    NuevaPlantillaFacturas -> libro nuevo con la plantilla de facturas
' =====================================================================
Option Explicit

' =====================================================================
'  MACROS PÚBLICAS
' =====================================================================
Public Sub EnlaceA3()
    If Not ConfiguracionLista() Then
        If MsgBox("Este libro todavía no tiene las hojas de configuración (EMPRESAS, IVA y PERFILES)." & vbCrLf & vbCrLf & _
                  "¿Las creo ahora?", vbQuestion + vbYesNo, "Enlace contable a3") = vbNo Then Exit Sub
        PrepararLibro
    End If
    frmEnlaceA3.Show
End Sub

Public Sub PrepararLibro()
    On Error GoTo Fallo
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    CrearHojasConfiguracion
    If Not ExisteHoja(HOJA_PLANT_DIARIO) Then CrearPlantillaDiario ThisWorkbook
    If Not ExisteHoja(HOJA_PLANT_FACTURAS) Then CrearPlantillaFacturas ThisWorkbook
    If Not ExisteHoja(HOJA_INICIO) Then CrearHojaInicio
    ThisWorkbook.Worksheets(HOJA_INICIO).Move Before:=ThisWorkbook.Worksheets(1)
    ThisWorkbook.Worksheets(HOJA_INICIO).Activate
    Application.ScreenUpdating = True
    MsgBox "Libro preparado." & vbCrLf & vbCrLf & _
           "- INICIO: botones para abrir la herramienta y crear plantillas" & vbCrLf & _
           "- EMPRESAS: código a3, dígitos del plan y cuentas de cada empresa" & vbCrLf & _
           "- IVA: cuentas de IVA repercutido por empresa y tipo" & vbCrLf & _
           "- PERFILES: cómo leer el libro de facturas de cada programa" & vbCrLf & _
           "- PLANTILLA_DIARIO / PLANTILLA_FACTURAS: formatos de entrada" & vbCrLf & vbCrLf & _
           IIf(EsLibroConMacros(), "Recuerda guardar el libro.", _
               "ATENCIÓN: este libro todavía no está guardado como .xlsm. Guárdalo ahora como " & _
               "'Libro de Excel habilitado para macros' o perderás las macros."), vbInformation, "Enlace contable a3"
    Exit Sub
Fallo:
    Application.ScreenUpdating = True
    MsgBox "No se ha podido preparar el libro: " & Err.Description, vbExclamation, "Enlace contable a3"
End Sub

Public Sub NuevaPlantillaDiario()
    Dim wb As Object
    Set wb = Application.Workbooks.Add
    CrearPlantillaDiario wb
    BorrarHojasSobrantes wb, HOJA_PLANT_DIARIO
    wb.Worksheets(1).Name = "Diario"
    MsgBox "Plantilla de diario creada en un libro nuevo." & vbCrLf & _
           "Rellénala (datos desde la fila 2) y guárdala donde quieras.", vbInformation, "Enlace contable a3"
End Sub

Public Sub NuevaPlantillaFacturas()
    Dim wb As Object
    Set wb = Application.Workbooks.Add
    CrearPlantillaFacturas wb
    BorrarHojasSobrantes wb, HOJA_PLANT_FACTURAS
    wb.Worksheets(1).Name = "Facturas emitidas"
    MsgBox "Plantilla de facturas emitidas creada en un libro nuevo." & vbCrLf & _
           "Una fila por factura y tipo de IVA. Usa el perfil GENERAL al generar.", vbInformation, "Enlace contable a3"
End Sub

Private Sub BorrarHojasSobrantes(ByVal wb As Object, ByVal conservar As String)
    Dim i As Long, alertas As Boolean
    alertas = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For i = wb.Worksheets.Count To 1 Step -1
        If wb.Worksheets(i).Name <> conservar And wb.Worksheets.Count > 1 Then wb.Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = alertas
End Sub

' =====================================================================
'  ORQUESTACIÓN DE UNA GENERACIÓN
'  soloAnalizar = True -> no escribe nada, solo devuelve el informe
' =====================================================================
Public Function EjecutarEnlace(ByRef p As TParametros, ByRef fil As TFiltros, ByVal soloAnalizar As Boolean, _
                               ByRef informe As String, ByRef rutaDatFinal As String) As Boolean
    Dim emp As TEmpresa, per As TPerfil, res As TResumen, datos As Variant, msg As String
    Dim rutaDat As String, rutaCtrl As String, copia As String, ok As Boolean, avisoFinal As String

    informe = ""
    rutaDatFinal = ""
    On Error GoTo Fallo
    Application.Cursor = 2                         ' xlWait

    ' --- empresa ----------------------------------------------------------
    p.Empresa = CodigoEmpresaA3(p.Empresa)
    If Not SoloDigitos(p.Empresa) Or p.Empresa = "00000" Then
        informe = "El código de empresa debe ser un número de hasta 5 dígitos."
        GoTo Salir
    End If
    If p.Digitos < 6 Or p.Digitos > 12 Then
        informe = "Los dígitos del plan de cuentas deben estar entre 6 y 12."
        GoTo Salir
    End If
    CargarEmpresa p.Empresa, emp
    emp.Digitos = p.Digitos

    ' --- origen -----------------------------------------------------------
    If p.Modo = MODO_EMITIDAS Then
        If Not emp.Existe Then
            informe = "La empresa " & p.Empresa & " no está configurada en la hoja EMPRESAS." & vbCrLf & _
                      "Pulsa ""Configurar empresa"" para darla de alta con cuentas por defecto."
            GoTo Salir
        End If
        If Not CargarPerfil(p.Perfil, per, msg) Then
            informe = msg
            GoTo Salir
        End If
        CargarIVA emp.Codigo
        If gNIva = 0 Then
            informe = "La empresa " & p.Empresa & " no tiene ningún tipo de IVA en la hoja IVA."
            GoTo Salir
        End If
    End If
    If Not ObtenerDatos(p, per.Separador, datos, msg) Then
        informe = msg
        GoTo Salir
    End If

    ' --- proceso ------------------------------------------------------------
    If p.Modo = MODO_DIARIO Then
        ok = DiarioProcesar(emp, fil, datos, res, msg)
    Else
        ok = FacturasProcesar(emp, per, fil, datos, res, msg)
    End If
    If Not ok Then
        informe = msg
        GoTo Salir
    End If

    rutaDat = UnirRuta(p.CarpetaSalida, A3_NOMBRE_FICHERO)
    informe = TextoInforme(p, emp, fil, res, IIf(soloAnalizar, "", rutaDat))
    If soloAnalizar Then
        EjecutarEnlace = True
        GoTo Salir
    End If

    ' --- escritura -----------------------------------------------------------
    If gNDat = 0 Then
        informe = "No hay nada que exportar con estos datos y filtros." & vbCrLf & vbCrLf & informe
        GoTo Salir
    End If
    Application.Cursor = -4143                     ' xlDefault
    If res.UnidadesExcluidas > 0 Then
        If MsgBox("Hay " & res.UnidadesExcluidas & IIf(p.Modo = MODO_DIARIO, " asientos", " documentos") & _
                  " excluidos por incidencias (se detallan en el Excel de control)." & vbCrLf & vbCrLf & _
                  "¿Genero el SUENLACE.DAT con el resto?", vbQuestion + vbYesNo, "Enlace contable a3") = vbNo Then
            informe = "Generación cancelada. Revisa las incidencias." & vbCrLf & vbCrLf & informe
            GoTo Salir
        End If
    End If
    Application.Cursor = 2
    If Trim$(p.CarpetaSalida) = "" Then
        informe = "Indica la carpeta de salida."
        GoTo Salir
    End If
    If Not CrearCarpeta(p.CarpetaSalida) Then
        informe = "No se puede crear o acceder a la carpeta de salida:" & vbCrLf & p.CarpetaSalida
        GoTo Salir
    End If
    If p.CopiaSeguridad Then copia = CopiaSeguridadDat(rutaDat)
    EscribirDat rutaDat
    rutaDatFinal = rutaDat

    ' A partir de aquí el fichero ya está escrito: un fallo solo es un aviso
    On Error Resume Next
    If p.GenerarControl Then
        rutaCtrl = UnirRuta(p.CarpetaSalida, "CONTROL_SUENLACE_" & emp.Codigo & "_" & MarcaTiempo() & ".xlsx")
        Err.Clear
        CrearControl p, emp, fil, res, rutaDat, rutaCtrl
        If Err.Number <> 0 Then
            avisoFinal = avisoFinal & "AVISO: no se ha podido crear o guardar el Excel de control: " & Err.Description & vbCrLf
            rutaCtrl = ""
        End If
    End If
    Err.Clear
    emp.CarpetaSalida = p.CarpetaSalida
    If p.Modo = MODO_EMITIDAS Then emp.Perfil = p.Perfil
    GuardarUsoEmpresa emp
    If Err.Number <> 0 Then avisoFinal = avisoFinal & "AVISO: no se ha podido actualizar la hoja EMPRESAS: " & Err.Description & vbCrLf
    Err.Clear
    GuardarPreferencia "A3_ULTIMA_EMPRESA", emp.Codigo
    GuardarPreferencia "A3_ULTIMO_MODO", CStr(p.Modo)
    GuardarLibroHerramienta
    Err.Clear
    On Error GoTo Fallo

    informe = avisoFinal & "SUENLACE.DAT GENERADO" & vbCrLf & rutaDat & vbCrLf & _
              IIf(copia <> "", "Copia del anterior: " & NombreDeRuta(copia) & vbCrLf, "") & _
              IIf(rutaCtrl <> "", "Excel de control: " & NombreDeRuta(rutaCtrl) & vbCrLf, "") & vbCrLf & informe
    EjecutarEnlace = True

Salir:
    Application.Cursor = -4143
    Exit Function
Fallo:
    Application.Cursor = -4143
    informe = "ERROR: " & Err.Description & IIf(Err.Source <> "", " (" & Err.Source & ")", "")
    EjecutarEnlace = False
End Function

' Carga los datos del origen como matriz (1..filas, 1..columnas)
Public Function ObtenerDatos(ByRef p As TParametros, ByVal separador As String, ByRef datos As Variant, ByRef msg As String) As Boolean
    Dim wb As Object, ws As Object
    If p.OrigenTipo = ORIGEN_CSV Then
        If p.OrigenRuta = "" Or Not ExisteFichero(p.OrigenRuta) Then
            msg = "No encuentro el fichero de origen:" & vbCrLf & p.OrigenRuta
            Exit Function
        End If
        If separador = "" Then separador = "auto"
        datos = ParsearCSV(LeerFicheroTexto(p.OrigenRuta), separador)
        p.OrigenDescripcion = p.OrigenRuta
    Else
        On Error Resume Next
        Set wb = Application.Workbooks(p.OrigenLibro)
        If Not wb Is Nothing Then Set ws = wb.Worksheets(p.OrigenHoja)
        On Error GoTo 0
        If ws Is Nothing Then
            msg = "No encuentro la hoja """ & p.OrigenHoja & """ del libro """ & p.OrigenLibro & """." & vbCrLf & _
                  "Comprueba que el libro sigue abierto."
            Exit Function
        End If
        datos = LeerHojaComoMatriz(ws)
        p.OrigenDescripcion = wb.FullName & "  [hoja " & ws.Name & "]"
        On Error Resume Next
        If ws.FilterMode Then p.OrigenDescripcion = p.OrigenDescripcion & _
            "  (ATENCIÓN: la hoja tiene un filtro activo; se leen también las filas ocultas)"
        On Error GoTo 0
    End If
    If ColumnasMatriz(datos) = 0 Then
        msg = "El origen está vacío."
        Exit Function
    End If
    If UBound(datos, 1) < 2 Then
        msg = "El origen solo tiene una fila: faltan los datos."
        Exit Function
    End If
    ObtenerDatos = True
End Function

' Lee el bloque usado de una hoja empezando en A1
Public Function LeerHojaComoMatriz(ByVal ws As Object) As Variant
    Dim c As Object, ultF As Long, ultC As Long
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=-4123, LookAt:=2, SearchOrder:=1, SearchDirection:=2)
    If c Is Nothing Then
        LeerHojaComoMatriz = ComoMatriz(Empty)
        Exit Function
    End If
    ultF = c.Row
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=-4123, LookAt:=2, SearchOrder:=2, SearchDirection:=2)
    ultC = c.Column
    LeerHojaComoMatriz = ComoMatriz(ws.Range(ws.Cells(1, 1), ws.Cells(ultF, ultC)).Value)
End Function

' Informe de texto para la ventana
Public Function TextoInforme(ByRef p As TParametros, ByRef emp As TEmpresa, ByRef fil As TFiltros, _
                             ByRef res As TResumen, ByVal rutaDat As String) As String
    Dim s As String, i As Long, n As Long, dif As Currency
    s = "EMPRESA " & emp.Codigo & IIf(emp.Nombre <> "", " - " & emp.Nombre, "") & vbCrLf
    If p.Modo = MODO_DIARIO Then
        s = s & "Diario (asientos sin IVA) · plan de " & emp.Digitos & " dígitos" & vbCrLf
    Else
        s = s & "Facturas emitidas · perfil " & p.Perfil & " · plan de " & emp.Digitos & " dígitos" & vbCrLf
    End If
    If p.OrigenDescripcion <> "" Then s = s & "Origen: " & p.OrigenDescripcion & vbCrLf
    s = s & "Filtros: " & TextoFiltros(p, fil) & vbCrLf
    If res.HayFechas Then s = s & "Fechas exportadas: " & FechaTexto(res.FechaMin) & " a " & FechaTexto(res.FechaMax) & vbCrLf
    s = s & String$(60, "-") & vbCrLf
    If p.Modo = MODO_DIARIO Then
        s = s & "Asientos: " & res.UnidadesLeidas & " en el filtro  ·  " & res.UnidadesExportadas & " exportados  ·  " & _
            res.UnidadesExcluidas & " excluidos" & vbCrLf
        s = s & "Apuntes en el SUENLACE: " & res.LineasDat & IIf(res.FilasFiltradas > 0, "  (" & res.FilasFiltradas & " fuera del filtro)", "") & vbCrLf
        dif = res.Debe - res.Haber
        s = s & "Debe " & ImporteTexto(res.Debe) & "  ·  Haber " & ImporteTexto(res.Haber) & _
            IIf(dif = 0, "  ·  CUADRA", "  ·  DIFERENCIA " & ImporteTexto(dif)) & vbCrLf
    Else
        s = s & "Documentos: " & res.UnidadesLeidas & " en el filtro  ·  " & res.UnidadesExportadas & " exportados (" & _
            res.NumFacturas & " facturas, " & res.NumTickets & " tickets, " & res.NumAbonos & " abonos)  ·  " & _
            res.UnidadesExcluidas & " excluidos" & vbCrLf
        s = s & "Líneas del SUENLACE: " & res.LineasDat & IIf(res.FilasFiltradas > 0, "  (" & res.FilasFiltradas & " filas fuera del filtro)", "") & vbCrLf
        s = s & "Base " & ImporteTexto(res.BaseImp) & "  ·  IVA " & ImporteTexto(res.Cuota)
        If res.CuotaRE <> 0 Then s = s & "  ·  Recargo " & ImporteTexto(res.CuotaRE)
        If res.CuotaRet <> 0 Then s = s & "  ·  Retención " & ImporteTexto(res.CuotaRet)
        s = s & "  ·  Total " & ImporteTexto(res.Total) & vbCrLf
        dif = res.TotalOrigen - res.TotalExcluido - res.AjusteSignoAbonos - res.Total
        s = s & "Cuadre con el origen: " & IIf(dif = 0, "OK", "NO CUADRA (diferencia " & ImporteTexto(dif) & ")") & vbCrLf
    End If
    s = s & String$(60, "-") & vbCrLf
    If gNInc = 0 Then
        s = s & "Sin incidencias." & vbCrLf
    Else
        s = s & "Incidencias: " & res.Excluidos & " excluidos, " & res.Avisos & " avisos" & vbCrLf
        n = 0
        For i = 1 To gNInc
            If gInc(i).Gravedad <> INC_INFO Then
                n = n + 1
                If n <= 12 Then
                    s = s & "  " & IIf(gInc(i).Gravedad = INC_EXCLUIDO, "[EXCLUIDO] ", "[aviso] ") & _
                        IIf(gInc(i).Referencia <> "", gInc(i).Referencia & ": ", "") & gInc(i).Texto & vbCrLf
                End If
            End If
        Next i
        If n > 12 Then s = s & "  ... y " & (n - 12) & " más (ver Excel de control)" & vbCrLf
    End If
    TextoInforme = s
End Function

' =====================================================================
'  PREFERENCIAS (nombres definidos ocultos en el propio libro)
' =====================================================================
Public Sub GuardarPreferencia(ByVal nombre As String, ByVal valor As String)
    On Error Resume Next
    ThisWorkbook.Names(nombre).Delete
    ThisWorkbook.Names.Add Name:=nombre, RefersTo:="=""" & Replace(valor, """", """""") & """", Visible:=False
    On Error GoTo 0
End Sub

Public Function LeerPreferencia(ByVal nombre As String, ByVal porDefecto As String) As String
    Dim v As String
    LeerPreferencia = porDefecto
    On Error Resume Next
    v = ThisWorkbook.Names(nombre).RefersTo
    If Err.Number = 0 And Len(v) >= 3 Then
        If Left$(v, 2) = "=""" And Right$(v, 1) = """" Then
            LeerPreferencia = Replace(Mid$(v, 3, Len(v) - 3), """""", """")
        End If
    End If
    On Error GoTo 0
End Function

Private Sub GuardarLibroHerramienta()
    On Error Resume Next
    If EsLibroConMacros() And Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    On Error GoTo 0
End Sub

' ¿Está la herramienta guardada ya como libro con macros (.xlsm / .xlsb)?
Public Function EsLibroConMacros() As Boolean
    On Error Resume Next
    EsLibroConMacros = (ThisWorkbook.Path <> "" And (ThisWorkbook.FileFormat = 52 Or ThisWorkbook.FileFormat = 50))
    On Error GoTo 0
End Function

' =====================================================================
'  HOJA INICIO
' =====================================================================
Private Sub CrearHojaInicio()
    Dim ws As Object, logo As String, sh As Object
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = HOJA_INICIO
    ws.Tab.Color = COLOR_ACP
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Interior.Color = COLOR_BLANCO
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:L").ColumnWidth = 11

    With ws.Range("A1:L4")
        .Interior.Color = COLOR_ACP
    End With
    With ws.Range("C2")
        .Value = "ENLACE CONTABLE A3"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = COLOR_BLANCO
    End With
    With ws.Range("C3")
        .Value = "ACP Asociados  ·  Córdoba   —   diarios y libros de facturas para a3ECO / a3CON (SUENLACE.DAT)"
        .Font.Size = 11
        .Font.Color = COLOR_ACP_MEDIO
    End With
    ws.Rows(2).RowHeight = 34
    ws.Rows(3).RowHeight = 20

    ' logo (opcional): logo_acp.png junto al libro
    logo = UnirRuta(ThisWorkbook.Path, "logo_acp.png")
    If ExisteFichero(logo) Then
        On Error Resume Next
        Set sh = ws.Shapes.AddPicture(logo, False, True, ws.Range("J1").Left, 4, 60, 43)
        On Error GoTo 0
    End If

    BotonInicio ws, "btnA3Abrir", "ABRIR ENLACE A3", "EnlaceA3", ws.Range("B7").Left, ws.Range("B7").Top, 260, 54, COLOR_ACP, 16
    BotonInicio ws, "btnA3PlantDiario", "Nueva plantilla de DIARIO", "NuevaPlantillaDiario", _
                ws.Range("B12").Left, ws.Range("B12").Top, 200, 34, COLOR_ACP_OSCURO, 11
    BotonInicio ws, "btnA3PlantFact", "Nueva plantilla de FACTURAS", "NuevaPlantillaFacturas", _
                ws.Range("B12").Left + 210, ws.Range("B12").Top, 200, 34, COLOR_ACP_OSCURO, 11

    ws.Range("B16").Value = "CÓMO SE USA"
    ws.Range("B16").Font.Bold = True
    ws.Range("B16").Font.Color = COLOR_ACP
    ws.Range("B17").Value = "1.  Abre el libro con los datos (diario o libro de facturas) o tenlo a mano en CSV."
    ws.Range("B18").Value = "2.  Pulsa ABRIR ENLACE A3: código de empresa, tipo de enlace, dígitos del plan y origen de los datos."
    ws.Range("B19").Value = "3.  Ajusta los filtros (fechas, documentos, series...) y pulsa ANALIZAR para ver el resultado sin escribir nada."
    ws.Range("B20").Value = "4.  Pulsa GENERAR: se crea SUENLACE.DAT y un Excel de control con el resumen y las incidencias."
    ws.Range("B21").Value = "5.  En a3: Utilidades > Importar/Exportar > Enlace contable. Chequea (CH) antes de enlazar."
    ws.Range("B23").Value = "CONFIGURACIÓN"
    ws.Range("B23").Font.Bold = True
    ws.Range("B23").Font.Color = COLOR_ACP
    ws.Range("B24").Value = "EMPRESAS: una fila por empresa (código a3, dígitos, cuentas).   IVA: cuentas de IVA por tipo.   " & _
                            "PERFILES: columnas de cada programa de facturación."
    ws.Range("B25").Value = "Consejo: haz copia de seguridad de la empresa en a3 antes de importar, y prueba primero en una empresa de pruebas."
    ws.Range("B25").Font.Italic = True
    ws.Range("B17:B25").Font.Size = 10
    ws.Range("B24:B25").Font.Color = COLOR_GRIS
    ws.Range("B27").Value = "Versión " & A3_VERSION & "  ·  formato SUENLACE 4 (254 posiciones)"
    ws.Range("B27").Font.Size = 8
    ws.Range("B27").Font.Color = COLOR_GRIS
    ws.Range("A1").Select
End Sub

Private Sub BotonInicio(ByVal ws As Object, ByVal nombre As String, ByVal texto As String, ByVal macro As String, _
                        ByVal x As Double, ByVal y As Double, ByVal ancho As Double, ByVal alto As Double, _
                        ByVal color As Long, ByVal tam As Double)
    Dim sh As Object
    Set sh = ws.Shapes.AddShape(5, x, y, ancho, alto)          ' 5 = msoShapeRoundedRectangle
    sh.Name = nombre
    sh.Fill.ForeColor.RGB = color
    sh.Line.Visible = False
    With sh.TextFrame2
        .TextRange.Text = texto
        .TextRange.Font.Size = tam
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = COLOR_BLANCO
        .TextRange.ParagraphFormat.Alignment = 2                ' centrado
        .VerticalAnchor = 3                                     ' centrado vertical
    End With
    sh.OnAction = macro
End Sub

' =====================================================================
'  PLANTILLAS
' =====================================================================
Public Sub CrearPlantillaDiario(ByVal wb As Object)
    Dim ws As Object, cab As Variant, i As Long, ejemplo As Variant, r As Long
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = HOJA_PLANT_DIARIO
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    cab = Array("Fecha Devengo", "Asiento", "SUBCUENTA", "NOMBRE SUBCUENTA", "DESCRIPCION", "Documento", "Debe", "Haber")
    For i = 0 To UBound(cab)
        ws.Cells(1, i + 1).Value = cab(i)
    Next i
    EstiloCabecera ws.Range("A1:H1")
    ws.Rows(1).RowHeight = 28
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 9
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 32
    ws.Columns("E").ColumnWidth = 32
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G:H").ColumnWidth = 13
    ws.Columns("C").NumberFormat = "@"
    ws.Columns("F").NumberFormat = "@"
    ws.Columns("A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("G:H").NumberFormat = "#,##0.00"
    ejemplo = Array( _
        Array(DateSerial(Year(Date), 1, 1), 1, "570000000", "Caja, euros", "ASIENTO DE APERTURA", "", 1500, 0), _
        Array(DateSerial(Year(Date), 1, 1), 1, "572000001", "Banco cuenta principal", "ASIENTO DE APERTURA", "", 8500, 0), _
        Array(DateSerial(Year(Date), 1, 1), 1, "100000000", "Capital social", "ASIENTO DE APERTURA", "", 0, 10000))
    For r = 0 To UBound(ejemplo)
        For i = 0 To 7
            ws.Cells(r + 2, i + 1).Value = ejemplo(r)(i)
        Next i
    Next r
    ws.Range("A2:H4").Font.Color = COLOR_GRIS
    ws.Range("J1").Value = "INSTRUCCIONES"
    ws.Range("J1").Font.Bold = True
    ws.Range("J1").Font.Color = COLOR_ACP
    ws.Range("J2").Value = "· Una fila por apunte. Las filas grises son un ejemplo: bórralas."
    ws.Range("J3").Value = "· Todas las líneas de un asiento llevan el mismo nº de asiento y la misma fecha."
    ws.Range("J4").Value = "· Cada asiento debe cuadrar (Debe = Haber)."
    ws.Range("J5").Value = "· Subcuenta con todos los dígitos del plan (o con punto: 572.1)."
    ws.Range("J6").Value = "· Importes en Debe o en Haber (el otro a 0 o vacío)."
    ws.Range("J7").Value = "· Textos de más de 30 caracteres se recortan; las tildes se quitan."
    ws.Range("J8").Value = "· Documento en formato texto (máx. 10 caracteres). Al copiar de otro libro usa 'Pegar valores'."
    ws.Range("J9").Value = "· Un asiento con cualquier error o de una sola línea se excluye entero y sale en el Excel de control."
    ws.Range("J2:J9").Font.Color = COLOR_GRIS
    ws.Activate
    ActiveWindow.DisplayGridlines = True
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Tab.Color = COLOR_ACP_MEDIO
End Sub

Public Sub CrearPlantillaFacturas(ByVal wb As Object)
    Dim ws As Object, cab As Variant, i As Long, ejemplo As Variant, r As Long, obligatorio As Variant
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = HOJA_PLANT_FACTURAS
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    cab = Array("Fecha", "Nº Factura", "Tipo", "Cliente", "NIF", "CP", "Base imponible", "% IVA", "Cuota IVA", "Total", _
                "Cuenta cliente", "Cuenta ventas", "Nombre cuenta ventas", "% Recargo", "Cuota recargo", _
                "% Retención", "Cuota retención", "Subtipo", "Anulada")
    obligatorio = Array(True, True, False, False, False, False, True, True, True, False, False, False, False, False, False, False, False, False, False)
    For i = 0 To UBound(cab)
        ws.Cells(1, i + 1).Value = cab(i)
    Next i
    EstiloCabecera ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(cab) + 1))
    For i = 0 To UBound(cab)
        If Not obligatorio(i) Then ws.Cells(1, i + 1).Interior.Color = COLOR_ACP_OSCURO
    Next i
    ws.Rows(1).RowHeight = 30
    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:F").ColumnWidth = 11
    ws.Columns("G:J").ColumnWidth = 12
    ws.Columns("K:M").ColumnWidth = 14
    ws.Columns("N:S").ColumnWidth = 10
    ws.Columns("A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("B").NumberFormat = "@"
    ws.Columns("E:F").NumberFormat = "@"
    ws.Columns("K:L").NumberFormat = "@"
    ws.Columns("R").NumberFormat = "@"
    ws.Range("G:G,I:J,O:O,Q:Q").NumberFormat = "#,##0.00"
    ejemplo = Array( _
        Array(DateSerial(Year(Date), 1, 15), "F-2026-001", "Factura", "Cliente Ejemplo SL", "B12345674", "14001", 100, 21, 21, 121), _
        Array(DateSerial(Year(Date), 1, 15), "F-2026-001", "Factura", "Cliente Ejemplo SL", "B12345674", "14001", 50, 10, 5, 55), _
        Array(DateSerial(Year(Date), 1, 20), "T-0001", "Ticket", "", "", "", 20, 10, 2, 22), _
        Array(DateSerial(Year(Date), 1, 31), "R-2026-001", "Abono", "Cliente Ejemplo SL", "B12345674", "14001", -50, 10, -5, -55))
    For r = 0 To UBound(ejemplo)
        For i = 0 To 9
            ws.Cells(r + 2, i + 1).Value = ejemplo(r)(i)
        Next i
    Next r
    ws.Range("A2:S5").Font.Color = COLOR_GRIS
    ws.Range("U1").Value = "INSTRUCCIONES"
    ws.Range("U1").Font.Bold = True
    ws.Range("U1").Font.Color = COLOR_ACP
    ws.Range("U2").Value = "· Una fila por factura y tipo de IVA (si una factura tiene 10% y 21%, van dos filas con el mismo nº)."
    ws.Range("U3").Value = "· Cabecera azul = obligatoria. Azul oscuro = opcional (si falta o está vacía se usa la hoja EMPRESAS / IVA)."
    ws.Range("U4").Value = "· Tipo: Factura, Ticket, Abono o Rectificativa. Abono: en negativo o en positivo (siempre resta). " & _
                           "Rectificativa: con su signo (negativa resta, positiva suma)."
    ws.Range("U5").Value = "· Total vacío = base + cuota + recargo - retención. Si se informa, debe cuadrar."
    ws.Range("U6").Value = "· Subtipo a3: 01 interior (por defecto), 02 exenta, 03 intracomunitaria, 06 exportación..."
    ws.Range("U7").Value = "· Anulada: cualquier dato (X, SI...) excluye el documento."
    ws.Range("U8").Value = "· Las filas grises son un ejemplo: bórralas."
    ws.Range("U9").Value = "· Nº de factura: a3 guarda 10 caracteres (si es más largo se quitan separadores y se deja el final)."
    ws.Range("U10").Value = "· Columna opcional 'Impreso' (01 = 347, 02 = 349 bienes, 11 = 349 servicios). Vacía: 02 si subtipo 03/04."
    ws.Range("U11").Value = "· NIF sin guiones ni espacios; CP de 5 cifras (se usan para el 347)."
    ws.Range("U2:U11").Font.Color = COLOR_GRIS
    ws.Activate
    ActiveWindow.DisplayGridlines = True
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Tab.Color = COLOR_ACP_MEDIO
End Sub
