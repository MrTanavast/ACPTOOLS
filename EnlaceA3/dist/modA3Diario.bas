Attribute VB_Name = "modA3Diario"
' =====================================================================
'  ENLACE CONTABLE A3  ·  ACP ASOCIADOS
'  Módulo DIARIO (asientos sin IVA, registro tipo 0)
'
'  Plantilla de entrada (fila 1 = cabeceras, datos desde la fila 2):
'     A Fecha | B Asiento | C Subcuenta | D Nombre subcuenta |
'     E Descripción | F Documento | G Debe | H Haber
'  Si las cabeceras tienen esos nombres, las columnas pueden ir en
'  cualquier orden; si no se reconocen se usan las columnas A-H.
'
'  Mejoras frente a la versión anterior:
'     - I/M/U calculado después de quitar las líneas omitidas (antes un
'       asiento podía quedarse sin "U" y a3 lo rechazaba).
'     - Las líneas de un mismo asiento se agrupan aunque no estén seguidas.
'     - Comprobación de cuadre Debe = Haber por asiento.
'     - Cuentas con notación de punto de a3 (572.1 -> 572000001).
'     - Una fila con Debe y Haber a la vez genera dos apuntes (antes se
'       perdía el Haber).
'     - Un asiento con errores se excluye entero (nunca a medias).
' =====================================================================
Option Explicit

Public Type TLinDia
    Fila As Long
    FechaOk As Boolean
    Fecha As Date
    ClaveFecha As String
    Asiento As String
    AsientoNum As Double
    AsientoEsNum As Boolean
    Cuenta As String
    NombreCuenta As String
    Descripcion As String
    Documento As String
    DH As String
    Importe As Currency
    ErrorTxt As String
    DentroFiltro As Boolean
    Grupo As Long
    Sig As Long
    Estado As String
End Type

Public Type TAsiento
    Asiento As String
    Fecha As Date
    FechaOk As Boolean
    Primera As Long
    Ultima As Long
    NLineas As Long
    NoConsecutivo As Boolean
End Type

Public Type TCtrlDia
    Fecha As Date
    Asiento As String
    Cuenta As String
    NombreCuenta As String
    Descripcion As String
    Documento As String
    DH As String
    Debe As Currency
    Haber As Currency
    LineaApunte As String
End Type

Public gLinDia() As TLinDia
Public gNLinDia As Long
Public gAsi() As TAsiento
Public gNAsi As Long
Public gCtrlDia() As TCtrlDia
Public gNCtrlDia As Long

Private Const DCOL_FECHA As Long = 1
Private Const DCOL_ASIENTO As Long = 2
Private Const DCOL_CUENTA As Long = 3
Private Const DCOL_NOMBRE As Long = 4
Private Const DCOL_DESCRIPCION As Long = 5
Private Const DCOL_DOCUMENTO As Long = 6
Private Const DCOL_DEBE As Long = 7
Private Const DCOL_HABER As Long = 8

' =====================================================================
'  PROCESO COMPLETO
' =====================================================================
Public Function DiarioProcesar(ByRef emp As TEmpresa, ByRef fil As TFiltros, ByRef datos As Variant, _
                               ByRef res As TResumen, ByRef msgError As String) As Boolean
    Dim col(1 To 8) As Long, filaInicio As Long, a As Long

    ResumenReiniciar res
    msgError = ""
    DatReiniciar
    IncReiniciar
    gNLinDia = 0: ReDim gLinDia(1 To 64)
    gNAsi = 0: ReDim gAsi(1 To 16)
    gNCtrlDia = 0: ReDim gCtrlDia(1 To 64)

    If ColumnasMatriz(datos) = 0 Then
        msgError = "El origen no tiene datos."
        Exit Function
    End If
    If emp.Digitos < 6 Or emp.Digitos > 12 Then
        msgError = "Los dígitos del plan de cuentas deben estar entre 6 y 12."
        Exit Function
    End If

    ResolverColumnasDiario datos, col, filaInicio
    LeerLineasDiario emp, datos, col, filaInicio
    AplicarFiltrosDiario fil, res
    AgruparAsientos
    For a = 1 To gNAsi
        ProcesarAsiento emp, fil, a, res
    Next a

    res.LineasDat = gNDat
    res.Avisos = IncContar(INC_AVISO)
    res.Excluidos = IncContar(INC_EXCLUIDO)
    DiarioProcesar = True
End Function

' Columnas por nombre de cabecera; si no se reconocen, A-H por posición.
Private Sub ResolverColumnasDiario(ByRef datos As Variant, ByRef col() As Long, ByRef filaInicio As Long)
    Dim ok As Boolean, i As Long, nCols As Long, prueba As Date
    nCols = ColumnasMatriz(datos)
    ' ¿La fila 1 ya son datos (sin cabecera)?
    prueba = LeerFecha(datos(1, 1), "DMA", ok)
    If ok Then
        For i = 1 To 8
            If i <= nCols Then col(i) = i Else col(i) = 0
        Next i
        filaInicio = 1
        IncAgregar False, 0, "", 1, INC_INFO, "La fila 1 no tiene cabeceras: se leen los datos desde la fila 1 (columnas A-H)", ""
        Exit Sub
    End If
    filaInicio = 2
    col(DCOL_FECHA) = BuscarColumna(datos, 1, "Fecha Devengo|Fecha|Fecha asiento|Fecha apunte|Fecha contable")
    col(DCOL_ASIENTO) = BuscarColumna(datos, 1, "Asiento|Nº Asiento|Numero asiento|Num asiento|N asiento")
    col(DCOL_CUENTA) = BuscarColumna(datos, 1, "Subcuenta|Cuenta|Codigo cuenta|Cod cuenta|Codigo subcuenta")
    col(DCOL_NOMBRE) = BuscarColumna(datos, 1, "Nombre subcuenta|Nombre cuenta|Descripcion cuenta|Descripcion subcuenta|Titulo cuenta")
    col(DCOL_DESCRIPCION) = BuscarColumna(datos, 1, "Descripcion|Concepto|Descripcion apunte|Descripcion asiento")
    col(DCOL_DOCUMENTO) = BuscarColumna(datos, 1, "Documento|Referencia|Doc|Ref documento|Referencia documento")
    col(DCOL_DEBE) = BuscarColumna(datos, 1, "Debe|Importe debe|Cargo")
    col(DCOL_HABER) = BuscarColumna(datos, 1, "Haber|Importe haber|Abono")
    If col(DCOL_FECHA) = 0 Or col(DCOL_ASIENTO) = 0 Or col(DCOL_CUENTA) = 0 Or col(DCOL_DEBE) = 0 Or col(DCOL_HABER) = 0 Then
        For i = 1 To 8
            If i <= nCols Then col(i) = i Else col(i) = 0
        Next i
        IncAgregar False, 0, "", 1, INC_INFO, "Cabeceras no reconocidas: se usan las columnas A-H de la plantilla " & _
            "(Fecha, Asiento, Subcuenta, Nombre, Descripción, Documento, Debe, Haber)", ""
    End If
End Sub

Private Function CeldaD(ByRef datos As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c = 0 Then
        CeldaD = Empty
    Else
        CeldaD = datos(r, c)
    End If
End Function

' =====================================================================
'  LECTURA
' =====================================================================
Private Sub LeerLineasDiario(ByRef emp As TEmpresa, ByRef datos As Variant, ByRef col() As Long, ByVal filaInicio As Long)
    Dim r As Long, ok As Boolean, errores As String, fecha As Date, fechaOk As Boolean
    Dim asiento As String, cuenta As String, cuentaNorm As String, motivo As String
    Dim debe As Currency, haber As Currency, nombre As String, descr As String, docu As String

    For r = filaInicio To UBound(datos, 1)
        If Not FilaVacia(datos, r) Then
            errores = ""
            fecha = LeerFecha(CeldaD(datos, r, col(DCOL_FECHA)), "DMA", fechaOk)
            If Not fechaOk Then errores = errores & "; fecha no válida [" & ValorTexto(CeldaD(datos, r, col(DCOL_FECHA))) & "]"
            asiento = ValorTexto(CeldaD(datos, r, col(DCOL_ASIENTO)))
            If asiento = "" Then errores = errores & "; número de asiento vacío"
            cuenta = ValorTexto(CeldaD(datos, r, col(DCOL_CUENTA)))
            cuentaNorm = NormalizarCuenta(cuenta, emp.Digitos, motivo)
            If cuentaNorm = "" Then errores = errores & "; " & motivo
            nombre = ValorTexto(CeldaD(datos, r, col(DCOL_NOMBRE)))
            descr = ValorTexto(CeldaD(datos, r, col(DCOL_DESCRIPCION)))
            docu = ValorTexto(CeldaD(datos, r, col(DCOL_DOCUMENTO)))
            debe = LeerImporte(CeldaD(datos, r, col(DCOL_DEBE)), "auto", ok)
            If Not ok Then errores = errores & "; Debe no numérico [" & ValorTexto(CeldaD(datos, r, col(DCOL_DEBE))) & "]"
            haber = LeerImporte(CeldaD(datos, r, col(DCOL_HABER)), "auto", ok)
            If Not ok Then errores = errores & "; Haber no numérico [" & ValorTexto(CeldaD(datos, r, col(DCOL_HABER))) & "]"

            If errores <> "" Then
                AgregarLineaDiario r, fechaOk, fecha, asiento, cuenta, nombre, descr, docu, "D", 0, Mid$(errores, 3)
                If Not fechaOk Then gLinDia(gNLinDia).ClaveFecha = "?" & UCase$(ValorTexto(CeldaD(datos, r, col(DCOL_FECHA))))
            ElseIf debe = 0 And haber = 0 Then
                IncAgregar fechaOk, fecha, asiento, r, INC_AVISO, "Debe y Haber a 0 en la cuenta " & cuentaNorm, "Línea omitida"
            Else
                If debe <> 0 And haber <> 0 Then
                    IncAgregar fechaOk, fecha, asiento, r, INC_AVISO, _
                        "La fila tiene Debe y Haber a la vez: se generan dos apuntes", "Incluido - verificar"
                End If
                If debe <> 0 Then AgregarLineaDiario r, fechaOk, fecha, asiento, cuentaNorm, nombre, descr, docu, "D", debe, ""
                If haber <> 0 Then AgregarLineaDiario r, fechaOk, fecha, asiento, cuentaNorm, nombre, descr, docu, "H", haber, ""
            End If
        End If
    Next r
End Sub

Private Sub AgregarLineaDiario(ByVal fila As Long, ByVal fechaOk As Boolean, ByVal fecha As Date, ByVal asiento As String, _
        ByVal cuenta As String, ByVal nombre As String, ByVal descr As String, ByVal docu As String, _
        ByVal dh As String, ByVal importe As Currency, ByVal errorTxt As String)
    Dim n As Long
    gNLinDia = gNLinDia + 1
    If gNLinDia > UBound(gLinDia) Then
        ReDim Preserve gLinDia(1 To UBound(gLinDia) * 2)
    End If
    n = gNLinDia
    gLinDia(n).Fila = fila
    gLinDia(n).FechaOk = fechaOk
    gLinDia(n).Fecha = fecha
    If fechaOk Then
        gLinDia(n).ClaveFecha = FechaA3(fecha)
    Else
        gLinDia(n).ClaveFecha = "?"
    End If
    gLinDia(n).Asiento = asiento
    gLinDia(n).AsientoEsNum = SoloDigitos(asiento)
    If gLinDia(n).AsientoEsNum Then gLinDia(n).AsientoNum = Val(asiento)
    gLinDia(n).Cuenta = cuenta
    gLinDia(n).NombreCuenta = nombre
    gLinDia(n).Descripcion = descr
    gLinDia(n).Documento = docu
    gLinDia(n).DH = dh
    gLinDia(n).Importe = importe
    gLinDia(n).ErrorTxt = errorTxt
    gLinDia(n).DentroFiltro = True
End Sub

' =====================================================================
'  FILTROS
' =====================================================================
Private Sub AplicarFiltrosDiario(ByRef fil As TFiltros, ByRef res As TResumen)
    Dim i As Long, dentro As Boolean
    res.FilasLeidas = gNLinDia
    For i = 1 To gNLinDia
        dentro = True
        If gLinDia(i).FechaOk Then
            If fil.UsarDesde Then
                If gLinDia(i).Fecha < fil.FechaDesde Then dentro = False
            End If
            If fil.UsarHasta Then
                If gLinDia(i).Fecha > fil.FechaHasta Then dentro = False
            End If
        End If
        If dentro Then dentro = AsientoEnRango(gLinDia(i).Asiento, fil.RefDesde, fil.RefHasta)
        gLinDia(i).DentroFiltro = dentro
        If Not dentro Then
            gLinDia(i).Estado = "Fuera del filtro"
            res.FilasFiltradas = res.FilasFiltradas + 1
        End If
    Next i
End Sub

Public Function AsientoEnRango(ByVal asiento As String, ByVal desde As String, ByVal hasta As String) As Boolean
    desde = Trim$(desde): hasta = Trim$(hasta)
    AsientoEnRango = True
    If desde <> "" Then
        If SoloDigitos(asiento) And SoloDigitos(desde) Then
            If Val(asiento) < Val(desde) Then AsientoEnRango = False
        ElseIf Normalizar(asiento) < Normalizar(desde) Then
            AsientoEnRango = False
        End If
    End If
    If hasta <> "" Then
        If SoloDigitos(asiento) And SoloDigitos(hasta) Then
            If Val(asiento) > Val(hasta) Then AsientoEnRango = False
        ElseIf Normalizar(asiento) > Normalizar(hasta) Then
            AsientoEnRango = False
        End If
    End If
End Function

' =====================================================================
'  AGRUPACIÓN: asiento + fecha, en orden de aparición
' =====================================================================
Private Sub AgruparAsientos()
    Dim i As Long, clave As String, idx As Long, ultimoGrupo As Long, fechaPrimera As Variant
    Dim indice As New Collection, fechas As New Collection, avisadoFechas As New Collection

    For i = 1 To gNLinDia
        If gLinDia(i).DentroFiltro Then
            clave = UCase$(gLinDia(i).Asiento) & "|" & gLinDia(i).ClaveFecha
            idx = 0
            On Error Resume Next
            idx = indice(clave)
            On Error GoTo 0
            If idx = 0 Then
                gNAsi = gNAsi + 1
                If gNAsi > UBound(gAsi) Then
                    ReDim Preserve gAsi(1 To UBound(gAsi) * 2)
                End If
                idx = gNAsi
                indice.Add idx, clave
                gAsi(idx).Asiento = gLinDia(i).Asiento
                gAsi(idx).Fecha = gLinDia(i).Fecha
                gAsi(idx).FechaOk = gLinDia(i).FechaOk
                gAsi(idx).Primera = i
                ' mismo número de asiento con otra fecha
                If gLinDia(i).FechaOk And gLinDia(i).Asiento <> "" Then
                    fechaPrimera = Empty
                    On Error Resume Next
                    fechaPrimera = fechas(UCase$(gLinDia(i).Asiento))
                    On Error GoTo 0
                    If IsEmpty(fechaPrimera) Then
                        fechas.Add gLinDia(i).Fecha, UCase$(gLinDia(i).Asiento)
                    ElseIf CDate(fechaPrimera) <> gLinDia(i).Fecha Then
                        If Not ClaveExiste(avisadoFechas, UCase$(gLinDia(i).Asiento)) Then
                            avisadoFechas.Add 1, UCase$(gLinDia(i).Asiento)
                            IncAgregar True, gLinDia(i).Fecha, gLinDia(i).Asiento, gLinDia(i).Fila, INC_AVISO, _
                                "El asiento " & gLinDia(i).Asiento & " tiene líneas con fechas distintas", _
                                "Se genera un asiento por cada fecha - verificar"
                        End If
                    End If
                End If
            Else
                If ultimoGrupo <> idx Then gAsi(idx).NoConsecutivo = True
                gLinDia(gAsi(idx).Ultima).Sig = i
            End If
            gAsi(idx).Ultima = i
            gAsi(idx).NLineas = gAsi(idx).NLineas + 1
            gLinDia(i).Grupo = idx
            gLinDia(i).Sig = 0
            ultimoGrupo = idx
        End If
    Next i
End Sub

Private Function ClaveExiste(ByVal c As Collection, ByVal clave As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c(clave)
    ClaveExiste = (Err.Number = 0)
    On Error GoTo 0
End Function

' =====================================================================
'  PROCESO DE UN ASIENTO
' =====================================================================
Private Sub ProcesarAsiento(ByRef emp As TEmpresa, ByRef fil As TFiltros, ByVal a As Long, ByRef res As TResumen)
    Dim i As Long, textoErr As String, nErr As Long, sumD As Currency, sumH As Currency
    Dim k As Long, lineaIMU As String, registro As String, primera As Long, excluir As Boolean

    primera = gAsi(a).Primera
    res.UnidadesLeidas = res.UnidadesLeidas + 1

    ' 1) errores de lectura: el asiento se excluye entero
    i = primera
    Do While i > 0
        If gLinDia(i).ErrorTxt <> "" Then
            nErr = nErr + 1
            If nErr <= 3 Then textoErr = textoErr & IIf(textoErr = "", "", " | ") & "fila " & gLinDia(i).Fila & ": " & gLinDia(i).ErrorTxt
        End If
        If gLinDia(i).DH = "D" Then sumD = sumD + gLinDia(i).Importe Else sumH = sumH + gLinDia(i).Importe
        i = gLinDia(i).Sig
    Loop
    If nErr > 3 Then textoErr = textoErr & " | y " & (nErr - 3) & " filas más"
    If nErr > 0 Then
        ExcluirAsiento a, "Datos no válidos (" & textoErr & ")", "Excluido - corregir y volver a generar", res
        Exit Sub
    End If

    ' 2) cuadre
    If sumD <> sumH Then
        If fil.ExcluirDescuadrados Then
            ExcluirAsiento a, "Asiento descuadrado: Debe " & ImporteTexto(sumD) & ", Haber " & ImporteTexto(sumH) & _
                " (diferencia " & ImporteTexto(sumD - sumH) & ")", "Excluido - corregir y volver a generar", res
            Exit Sub
        Else
            IncAgregar gAsi(a).FechaOk, gAsi(a).Fecha, gAsi(a).Asiento, gLinDia(primera).Fila, INC_AVISO, _
                "Asiento descuadrado: Debe " & ImporteTexto(sumD) & ", Haber " & ImporteTexto(sumH) & _
                " (diferencia " & ImporteTexto(sumD - sumH) & ")", "Incluido - a3 lo avisará en el chequeo"
        End If
    End If
    If gAsi(a).NLineas = 1 Then
        IncAgregar gAsi(a).FechaOk, gAsi(a).Fecha, gAsi(a).Asiento, gLinDia(primera).Fila, INC_AVISO, _
            "Asiento de una sola línea", "Incluido - verificar"
    End If
    If gAsi(a).NoConsecutivo Then
        IncAgregar gAsi(a).FechaOk, gAsi(a).Fecha, gAsi(a).Asiento, gLinDia(primera).Fila, INC_INFO, _
            "Las líneas del asiento no estaban seguidas en el origen", "Se han juntado en un solo asiento"
    End If

    ' 3) registros I / M / U
    k = 0
    i = primera
    Do While i > 0
        k = k + 1
        If k = 1 Then
            lineaIMU = "I"
        ElseIf k = gAsi(a).NLineas Then
            lineaIMU = "U"
        Else
            lineaIMU = "M"
        End If
        registro = RegistroApunte(emp.Codigo, gLinDia(i).Fecha, gLinDia(i).Cuenta, gLinDia(i).NombreCuenta, _
                                  gLinDia(i).DH, gLinDia(i).Documento, lineaIMU, gLinDia(i).Descripcion, gLinDia(i).Importe)
        DatAgregar registro
        gLinDia(i).Estado = "Exportada"

        gNCtrlDia = gNCtrlDia + 1
        If gNCtrlDia > UBound(gCtrlDia) Then
            ReDim Preserve gCtrlDia(1 To UBound(gCtrlDia) * 2)
        End If
        gCtrlDia(gNCtrlDia).Fecha = gLinDia(i).Fecha
        gCtrlDia(gNCtrlDia).Asiento = gLinDia(i).Asiento
        gCtrlDia(gNCtrlDia).Cuenta = gLinDia(i).Cuenta
        gCtrlDia(gNCtrlDia).NombreCuenta = gLinDia(i).NombreCuenta
        gCtrlDia(gNCtrlDia).Descripcion = gLinDia(i).Descripcion
        gCtrlDia(gNCtrlDia).Documento = gLinDia(i).Documento
        gCtrlDia(gNCtrlDia).DH = gLinDia(i).DH
        If gLinDia(i).DH = "D" Then
            gCtrlDia(gNCtrlDia).Debe = gLinDia(i).Importe
            res.Debe = res.Debe + gLinDia(i).Importe
        Else
            gCtrlDia(gNCtrlDia).Haber = gLinDia(i).Importe
            res.Haber = res.Haber + gLinDia(i).Importe
        End If
        gCtrlDia(gNCtrlDia).LineaApunte = lineaIMU
        i = gLinDia(i).Sig
    Loop

    res.UnidadesExportadas = res.UnidadesExportadas + 1
    If Not res.HayFechas Then
        res.FechaMin = gAsi(a).Fecha: res.FechaMax = gAsi(a).Fecha: res.HayFechas = True
    Else
        If gAsi(a).Fecha < res.FechaMin Then res.FechaMin = gAsi(a).Fecha
        If gAsi(a).Fecha > res.FechaMax Then res.FechaMax = gAsi(a).Fecha
    End If
End Sub

Private Sub ExcluirAsiento(ByVal a As Long, ByVal texto As String, ByVal tratamiento As String, ByRef res As TResumen)
    Dim i As Long
    IncAgregar gAsi(a).FechaOk, gAsi(a).Fecha, gAsi(a).Asiento, gLinDia(gAsi(a).Primera).Fila, INC_EXCLUIDO, texto, tratamiento
    i = gAsi(a).Primera
    Do While i > 0
        gLinDia(i).Estado = "Excluida"
        i = gLinDia(i).Sig
    Loop
    res.UnidadesExcluidas = res.UnidadesExcluidas + 1
End Sub
