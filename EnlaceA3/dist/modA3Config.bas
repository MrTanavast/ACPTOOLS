Attribute VB_Name = "modA3Config"
' =====================================================================
'  ENLACE CONTABLE A3  ·  ACP ASOCIADOS
'  Módulo CONFIGURACIÓN: hojas EMPRESAS, IVA y PERFILES del propio libro.
'
'  - EMPRESAS : una fila por empresa (código a3, dígitos del plan, cuentas...)
'  - IVA      : cuentas de IVA repercutido (y recargo) por empresa y tipo
'  - PERFILES : cómo se leen los libros de facturas de cada programa de
'               origen (una columna por programa). "GENERAL" es la plantilla
'               estándar de ACP; "CONTAPLUZ" es el CSV del programa del hotel.
' =====================================================================
Option Explicit

Public Const HOJA_INICIO As String = "INICIO"
Public Const HOJA_EMPRESAS As String = "EMPRESAS"
Public Const HOJA_IVA As String = "IVA"
Public Const HOJA_PERFILES As String = "PERFILES"
Public Const HOJA_PLANT_DIARIO As String = "PLANTILLA_DIARIO"
Public Const HOJA_PLANT_FACTURAS As String = "PLANTILLA_FACTURAS"

' --- Columnas de la hoja EMPRESAS -------------------------------------
Public Const EMP_CODIGO As Long = 1
Public Const EMP_NOMBRE As Long = 2
Public Const EMP_DIGITOS As Long = 3
Public Const EMP_PERFIL As Long = 4
Public Const EMP_CTA_CLIENTES As Long = 5
Public Const EMP_DESC_CLIENTES As Long = 6
Public Const EMP_CTA_VENTAS As Long = 7
Public Const EMP_DESC_VENTAS As Long = 8
Public Const EMP_CTA_VENTAS_ALT As Long = 9
Public Const EMP_DESC_VENTAS_ALT As Long = 10
Public Const EMP_CTA_RETENCION As Long = 11
Public Const EMP_CARPETA As Long = 12
Public Const EMP_ULTIMO_USO As Long = 13
Public Const EMP_NOTAS As Long = 14
Public Const EMP_NUM_COLUMNAS As Long = 14

' --- Campos de un libro de facturas (filas de la hoja PERFILES) --------
Public Const CAMPO_FECHA As Long = 1
Public Const CAMPO_DOCUMENTO As Long = 2
Public Const CAMPO_TIPO As Long = 3
Public Const CAMPO_CLIENTE As Long = 4
Public Const CAMPO_NIF As Long = 5
Public Const CAMPO_NIF2 As Long = 6
Public Const CAMPO_CP As Long = 7
Public Const CAMPO_BASE As Long = 8
Public Const CAMPO_PCT_IVA As Long = 9
Public Const CAMPO_CUOTA_IVA As Long = 10
Public Const CAMPO_TOTAL As Long = 11
Public Const CAMPO_ANULADA As Long = 12
Public Const CAMPO_SELECTOR As Long = 13
Public Const CAMPO_CTA_CLIENTE As Long = 14
Public Const CAMPO_DESC_CTA_CLIENTE As Long = 15
Public Const CAMPO_CTA_VENTAS As Long = 16
Public Const CAMPO_DESC_CTA_VENTAS As Long = 17
Public Const CAMPO_CTA_IVA As Long = 18
Public Const CAMPO_PCT_RE As Long = 19
Public Const CAMPO_CUOTA_RE As Long = 20
Public Const CAMPO_PCT_RET As Long = 21
Public Const CAMPO_CUOTA_RET As Long = 22
Public Const CAMPO_FECHA_OP As Long = 23
Public Const CAMPO_SUBTIPO As Long = 24
Public Const CAMPO_IMPRESO As Long = 25
Public Const NUM_CAMPOS As Long = 25

' --- Tabla de IVA de la empresa en curso ---------------------------------
Public gIvaPct() As Currency
Public gIvaCta() As String
Public gIvaPctRE() As Currency
Public gIvaCtaRE() As String
Public gNIva As Long

' =====================================================================
'  TABLA DE IVA (en memoria)
' =====================================================================
Public Sub IvaReiniciar()
    gNIva = 0
    ReDim gIvaPct(1 To 16)
    ReDim gIvaCta(1 To 16)
    ReDim gIvaPctRE(1 To 16)
    ReDim gIvaCtaRE(1 To 16)
End Sub

Public Sub IvaAgregar(ByVal pct As Currency, ByVal cta As String, ByVal pctRE As Currency, ByVal ctaRE As String)
    gNIva = gNIva + 1
    If gNIva > UBound(gIvaPct) Then
        ReDim Preserve gIvaPct(1 To gNIva * 2)
        ReDim Preserve gIvaCta(1 To gNIva * 2)
        ReDim Preserve gIvaPctRE(1 To gNIva * 2)
        ReDim Preserve gIvaCtaRE(1 To gNIva * 2)
    End If
    gIvaPct(gNIva) = pct
    gIvaCta(gNIva) = Trim$(cta)
    gIvaPctRE(gNIva) = pctRE
    gIvaCtaRE(gNIva) = Trim$(ctaRE)
End Sub

' Busca la cuenta de IVA para un tipo. Devuelve False si no está configurado.
Public Function IvaBuscar(ByVal pct As Currency, ByRef cta As String, ByRef pctRE As Currency, ByRef ctaRE As String) As Boolean
    Dim i As Long
    cta = "": ctaRE = "": pctRE = 0
    For i = 1 To gNIva
        If gIvaPct(i) = pct Then
            cta = gIvaCta(i)
            pctRE = gIvaPctRE(i)
            ctaRE = gIvaCtaRE(i)
            IvaBuscar = (cta <> "")
            Exit Function
        End If
    Next i
End Function

' =====================================================================
'  PERFILES
' =====================================================================

' Clave (fila de la hoja PERFILES) de cada campo
Public Function CampoClave(ByVal campo As Long) As String
    Select Case campo
        Case CAMPO_FECHA: CampoClave = "FECHA"
        Case CAMPO_DOCUMENTO: CampoClave = "DOCUMENTO"
        Case CAMPO_TIPO: CampoClave = "TIPO"
        Case CAMPO_CLIENTE: CampoClave = "CLIENTE"
        Case CAMPO_NIF: CampoClave = "NIF"
        Case CAMPO_NIF2: CampoClave = "NIF_ALTERNATIVO"
        Case CAMPO_CP: CampoClave = "CODIGO_POSTAL"
        Case CAMPO_BASE: CampoClave = "BASE"
        Case CAMPO_PCT_IVA: CampoClave = "PCT_IVA"
        Case CAMPO_CUOTA_IVA: CampoClave = "CUOTA_IVA"
        Case CAMPO_TOTAL: CampoClave = "TOTAL"
        Case CAMPO_ANULADA: CampoClave = "ANULADA"
        Case CAMPO_SELECTOR: CampoClave = "SELECTOR_VENTAS"
        Case CAMPO_CTA_CLIENTE: CampoClave = "CTA_CLIENTE"
        Case CAMPO_DESC_CTA_CLIENTE: CampoClave = "DESC_CTA_CLIENTE"
        Case CAMPO_CTA_VENTAS: CampoClave = "CTA_VENTAS"
        Case CAMPO_DESC_CTA_VENTAS: CampoClave = "DESC_CTA_VENTAS"
        Case CAMPO_CTA_IVA: CampoClave = "CTA_IVA"
        Case CAMPO_PCT_RE: CampoClave = "PCT_RECARGO"
        Case CAMPO_CUOTA_RE: CampoClave = "CUOTA_RECARGO"
        Case CAMPO_PCT_RET: CampoClave = "PCT_RETENCION"
        Case CAMPO_CUOTA_RET: CampoClave = "CUOTA_RETENCION"
        Case CAMPO_FECHA_OP: CampoClave = "FECHA_OPERACION"
        Case CAMPO_SUBTIPO: CampoClave = "SUBTIPO"
        Case CAMPO_IMPRESO: CampoClave = "IMPRESO"
    End Select
End Function

Public Function CampoDescripcion(ByVal campo As Long) As String
    Select Case campo
        Case CAMPO_FECHA: CampoDescripcion = "Fecha de la factura (obligatorio)"
        Case CAMPO_DOCUMENTO: CampoDescripcion = "Número de factura (obligatorio)"
        Case CAMPO_TIPO: CampoDescripcion = "Tipo de documento: factura, ticket o abono"
        Case CAMPO_CLIENTE: CampoDescripcion = "Nombre del cliente"
        Case CAMPO_NIF: CampoDescripcion = "NIF del cliente"
        Case CAMPO_NIF2: CampoDescripcion = "Segundo campo de NIF (se usa si el primero está vacío)"
        Case CAMPO_CP: CampoDescripcion = "Código postal del cliente"
        Case CAMPO_BASE: CampoDescripcion = "Base imponible (obligatorio)"
        Case CAMPO_PCT_IVA: CampoDescripcion = "% de IVA (obligatorio)"
        Case CAMPO_CUOTA_IVA: CampoDescripcion = "Cuota de IVA (obligatorio)"
        Case CAMPO_TOTAL: CampoDescripcion = "Total de la línea (si falta: base + cuotas - retención)"
        Case CAMPO_ANULADA: CampoDescripcion = "Marca de anulada (con dato = documento anulado)"
        Case CAMPO_SELECTOR: CampoDescripcion = "Columna que, si tiene dato, manda a la cuenta de ventas alternativa"
        Case CAMPO_CTA_CLIENTE: CampoDescripcion = "Cuenta del cliente (si falta: la de EMPRESAS)"
        Case CAMPO_DESC_CTA_CLIENTE: CampoDescripcion = "Nombre de la cuenta del cliente"
        Case CAMPO_CTA_VENTAS: CampoDescripcion = "Cuenta de ventas (si falta: la de EMPRESAS)"
        Case CAMPO_DESC_CTA_VENTAS: CampoDescripcion = "Nombre de la cuenta de ventas"
        Case CAMPO_CTA_IVA: CampoDescripcion = "Cuenta de IVA (si falta: la de la hoja IVA)"
        Case CAMPO_PCT_RE: CampoDescripcion = "% de recargo de equivalencia"
        Case CAMPO_CUOTA_RE: CampoDescripcion = "Cuota de recargo de equivalencia"
        Case CAMPO_PCT_RET: CampoDescripcion = "% de retención"
        Case CAMPO_CUOTA_RET: CampoDescripcion = "Cuota de retención"
        Case CAMPO_FECHA_OP: CampoDescripcion = "Fecha de operación (si falta: la de la factura)"
        Case CAMPO_SUBTIPO: CampoDescripcion = "Subtipo a3 (01 interior, 02 exenta, 03 intracomunitaria, 06 exportación, 08 ISP / no sujeta, 09 exenta con derecho)"
        Case CAMPO_IMPRESO: CampoDescripcion = "Impreso a3 (01 = 347; 02 = 349 bienes; 11 = 349 servicios). Vacío: 02 si subtipo 03/04, si no 01"
    End Select
End Function

' Nombre de columna configurado para un campo (con alternativas "a|b").
' Si empieza por "?" la columna es opcional: si no está, no pasa nada.
Public Function PerfilCampo(ByRef per As TPerfil, ByVal campo As Long) As String
    Dim p() As String
    p = Split(per.Campos, vbTab)
    If campo - 1 <= UBound(p) Then PerfilCampo = Trim$(p(campo - 1))
End Function

Public Function CampoObligatorio(ByVal campo As Long) As Boolean
    Select Case campo
        Case CAMPO_FECHA, CAMPO_DOCUMENTO, CAMPO_BASE, CAMPO_PCT_IVA, CAMPO_CUOTA_IVA
            CampoObligatorio = True
    End Select
End Function

' Perfiles de fábrica. Se usan para crear la hoja PERFILES y en las pruebas.
Public Function PerfilPredefinido(ByVal nombre As String, ByRef per As TPerfil) As Boolean
    Dim c(1 To NUM_CAMPOS) As String, i As Long
    Select Case UCase$(Trim$(nombre))
        Case "GENERAL"
            c(CAMPO_FECHA) = "Fecha|Fecha factura"
            c(CAMPO_DOCUMENTO) = "Nº Factura|Numero factura|Factura|Documento"
            c(CAMPO_TIPO) = "?Tipo|Tipo documento"
            c(CAMPO_CLIENTE) = "?Cliente|Nombre cliente"
            c(CAMPO_NIF) = "?NIF|CIF|NIF cliente"
            c(CAMPO_NIF2) = ""
            c(CAMPO_CP) = "?CP|Código postal"
            c(CAMPO_BASE) = "Base imponible|Base"
            c(CAMPO_PCT_IVA) = "% IVA|Tipo IVA|IVA %"
            c(CAMPO_CUOTA_IVA) = "Cuota IVA|Cuota"
            c(CAMPO_TOTAL) = "?Total|Total factura"
            c(CAMPO_ANULADA) = "?Anulada"
            c(CAMPO_SELECTOR) = ""
            c(CAMPO_CTA_CLIENTE) = "?Cuenta cliente"
            c(CAMPO_DESC_CTA_CLIENTE) = "?Nombre cuenta cliente"
            c(CAMPO_CTA_VENTAS) = "?Cuenta ventas"
            c(CAMPO_DESC_CTA_VENTAS) = "?Nombre cuenta ventas"
            c(CAMPO_CTA_IVA) = "?Cuenta IVA"
            c(CAMPO_PCT_RE) = "?% Recargo"
            c(CAMPO_CUOTA_RE) = "?Cuota recargo"
            c(CAMPO_PCT_RET) = "?% Retención"
            c(CAMPO_CUOTA_RET) = "?Cuota retención"
            c(CAMPO_FECHA_OP) = "?Fecha operación"
            c(CAMPO_SUBTIPO) = "?Subtipo"
            c(CAMPO_IMPRESO) = "?Impreso"
            per.Nombre = "GENERAL"
            per.Separador = "auto"
            per.SepDecimal = "auto"
            per.FormatoFecha = "DMA"
            per.ValoresAbono = "Abono;A;Nota de crédito;NC"
            per.ValoresTicket = "Ticket;T;Simplificada;S;Factura simplificada"
            per.ValoresRectificativa = "Rectificativa;R"
            per.NombreVacio = "Clientes varios"
            per.AbonoSiNegativo = True
            per.FilaCabecera = 1
        Case "CONTAPLUZ"
            c(CAMPO_FECHA) = "Date"
            c(CAMPO_DOCUMENTO) = "Document"
            c(CAMPO_TIPO) = "Type"
            c(CAMPO_CLIENTE) = "Holder"
            c(CAMPO_NIF) = "Vat Number"
            c(CAMPO_NIF2) = "NationalID"
            c(CAMPO_CP) = "Zip Code"
            c(CAMPO_BASE) = "Before Tax"
            c(CAMPO_PCT_IVA) = "Vat Percent"
            c(CAMPO_CUOTA_IVA) = "Tax"
            c(CAMPO_TOTAL) = "Total"
            c(CAMPO_ANULADA) = "Cancelled"
            c(CAMPO_SELECTOR) = "Center"
            per.Nombre = "CONTAPLUZ"
            per.Separador = ";"
            per.SepDecimal = "."
            per.FormatoFecha = "DMA"
            per.ValoresAbono = "Credit Note"
            per.ValoresTicket = "Receipt"
            per.ValoresRectificativa = ""
            per.NombreVacio = "Factura simple"
            per.AbonoSiNegativo = False
            per.FilaCabecera = 1
        Case Else
            Exit Function
    End Select
    per.Campos = ""
    For i = 1 To NUM_CAMPOS
        If i > 1 Then per.Campos = per.Campos & vbTab
        per.Campos = per.Campos & c(i)
    Next i
    PerfilPredefinido = True
End Function

' =====================================================================
'  LECTURA DE LA CONFIGURACIÓN DESDE LAS HOJAS
' =====================================================================
Public Function ExisteHoja(ByVal nombre As String, Optional ByVal libro As Object) As Boolean
    Dim ws As Object
    If libro Is Nothing Then Set libro = ThisWorkbook
    On Error Resume Next
    Set ws = libro.Worksheets(nombre)
    ExisteHoja = (Err.Number = 0 And Not ws Is Nothing)
    On Error GoTo 0
End Function

Public Function ConfiguracionLista() As Boolean
    ConfiguracionLista = ExisteHoja(HOJA_EMPRESAS) And ExisteHoja(HOJA_IVA) And ExisteHoja(HOJA_PERFILES)
End Function

Private Function UltimaFila(ByVal ws As Object, ByVal columna As Long) As Long
    UltimaFila = ws.Cells(ws.Rows.Count, columna).End(-4162).Row      ' -4162 = xlUp
End Function

' Busca una empresa por su código. Si no existe, devuelve emp.Existe = False
' con valores por defecto razonables.
Public Function CargarEmpresa(ByVal codigo As String, ByRef emp As TEmpresa) As Boolean
    Dim ws As Object, r As Long, n As Long
    emp.Codigo = CodigoEmpresaA3(codigo)
    emp.Nombre = "": emp.CtaClientes = "": emp.DescClientes = "": emp.CtaVentas = "": emp.DescVentas = ""
    emp.CtaVentasAlt = "": emp.DescVentasAlt = "": emp.CtaRetencion = "": emp.CarpetaSalida = ""
    emp.Existe = False: emp.Fila = 0
    emp.Digitos = 8
    emp.Perfil = "GENERAL"
    If Not ExisteHoja(HOJA_EMPRESAS) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(HOJA_EMPRESAS)
    n = UltimaFila(ws, EMP_CODIGO)
    For r = 2 To n
        If CodigoEmpresaA3(ValorTexto(ws.Cells(r, EMP_CODIGO).Value)) = emp.Codigo And ValorTexto(ws.Cells(r, EMP_CODIGO).Value) <> "" Then
            emp.Existe = True
            emp.Fila = r
            emp.Nombre = ValorTexto(ws.Cells(r, EMP_NOMBRE).Value)
            If IsNumeric(ws.Cells(r, EMP_DIGITOS).Value) And Not EsVacio(ws.Cells(r, EMP_DIGITOS).Value) Then
                emp.Digitos = CInt(ws.Cells(r, EMP_DIGITOS).Value)
            End If
            If ValorTexto(ws.Cells(r, EMP_PERFIL).Value) <> "" Then emp.Perfil = ValorTexto(ws.Cells(r, EMP_PERFIL).Value)
            emp.CtaClientes = ValorTexto(ws.Cells(r, EMP_CTA_CLIENTES).Value)
            emp.DescClientes = ValorTexto(ws.Cells(r, EMP_DESC_CLIENTES).Value)
            emp.CtaVentas = ValorTexto(ws.Cells(r, EMP_CTA_VENTAS).Value)
            emp.DescVentas = ValorTexto(ws.Cells(r, EMP_DESC_VENTAS).Value)
            emp.CtaVentasAlt = ValorTexto(ws.Cells(r, EMP_CTA_VENTAS_ALT).Value)
            emp.DescVentasAlt = ValorTexto(ws.Cells(r, EMP_DESC_VENTAS_ALT).Value)
            emp.CtaRetencion = ValorTexto(ws.Cells(r, EMP_CTA_RETENCION).Value)
            emp.CarpetaSalida = ValorTexto(ws.Cells(r, EMP_CARPETA).Value)
            CargarEmpresa = True
            Exit Function
        End If
    Next r
End Function

' Crea o actualiza la fila de la empresa (dígitos, perfil, carpeta y fecha de uso).
Public Sub GuardarUsoEmpresa(ByRef emp As TEmpresa)
    Dim ws As Object, r As Long
    If Not ExisteHoja(HOJA_EMPRESAS) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(HOJA_EMPRESAS)
    If emp.Fila > 0 Then
        r = emp.Fila
    Else
        r = UltimaFila(ws, EMP_CODIGO) + 1
        If r < 2 Then r = 2
        ws.Cells(r, EMP_CODIGO).NumberFormat = "@"
        ws.Cells(r, EMP_CODIGO).Value = emp.Codigo
        If emp.Nombre = "" Then emp.Nombre = "(pendiente de completar)"
        ws.Cells(r, EMP_NOMBRE).Value = emp.Nombre
        ws.Cells(r, EMP_PERFIL).Value = emp.Perfil
        emp.Fila = r
        emp.Existe = True
    End If
    ws.Cells(r, EMP_DIGITOS).Value = emp.Digitos
    If emp.CarpetaSalida <> "" Then ws.Cells(r, EMP_CARPETA).Value = emp.CarpetaSalida
    ws.Cells(r, EMP_ULTIMO_USO).Value = Now
    ws.Cells(r, EMP_ULTIMO_USO).NumberFormat = "dd/mm/yyyy hh:mm"
End Sub

' Da de alta una empresa nueva con cuentas por defecto según los dígitos del plan.
Public Sub AltaEmpresaPorDefecto(ByVal codigo As String, ByVal nombre As String, ByVal digitos As Integer, ByVal perfil As String)
    Dim ws As Object, wi As Object, r As Long, tipos As Variant, i As Long
    Set ws = ThisWorkbook.Worksheets(HOJA_EMPRESAS)
    Set wi = ThisWorkbook.Worksheets(HOJA_IVA)
    r = UltimaFila(ws, EMP_CODIGO) + 1
    If r < 2 Then r = 2
    ws.Cells(r, EMP_CODIGO).NumberFormat = "@"
    ws.Cells(r, EMP_CODIGO).Value = CodigoEmpresaA3(codigo)
    ws.Cells(r, EMP_NOMBRE).Value = nombre
    ws.Cells(r, EMP_DIGITOS).Value = digitos
    ws.Cells(r, EMP_PERFIL).Value = perfil
    ws.Range(ws.Cells(r, EMP_CTA_CLIENTES), ws.Cells(r, EMP_CTA_RETENCION)).NumberFormat = "@"
    ws.Cells(r, EMP_CTA_CLIENTES).Value = CuentaPatron("430", digitos, "")
    ws.Cells(r, EMP_DESC_CLIENTES).Value = "Clientes varios"
    ws.Cells(r, EMP_CTA_VENTAS).Value = CuentaPatron("700", digitos, "")
    ws.Cells(r, EMP_DESC_VENTAS).Value = "Ventas"
    ws.Cells(r, EMP_CTA_RETENCION).Value = CuentaPatron("473", digitos, "")
    ws.Cells(r, EMP_NOTAS).Value = "Alta automática: revisa las cuentas antes de importar"
    ws.Rows(r).Interior.Color = COLOR_AVISO_CLARO
    ' Tipos de IVA habituales
    tipos = Array(4, 10, 21)
    For i = LBound(tipos) To UBound(tipos)
        r = UltimaFila(wi, 1) + 1
        wi.Cells(r, 1).NumberFormat = "@"
        wi.Cells(r, 1).Value = CodigoEmpresaA3(codigo)
        wi.Cells(r, 2).Value = tipos(i)
        wi.Cells(r, 3).NumberFormat = "@"
        wi.Cells(r, 3).Value = CuentaPatron("477", digitos, Right$("00" & CStr(tipos(i)), 2))
        wi.Cells(r, 4).Value = 0
        wi.Cells(r, 6).Value = "Alta automática: revisar"
        wi.Rows(r).Interior.Color = COLOR_AVISO_CLARO
    Next i
End Sub

' Carga en memoria la tabla de IVA de una empresa (filas con su código;
' las filas con "*" valen para todas las empresas que no tengan ese tipo).
Public Sub CargarIVA(ByVal codigo As String)
    Dim ws As Object, r As Long, n As Long, cod As String, ok As Boolean
    Dim pct As Currency, pctRE As Currency, pasada As Long, cta As String, ctaRE As String, dummy As Currency
    Dim celda As Object
    IvaReiniciar
    If Not ExisteHoja(HOJA_IVA) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(HOJA_IVA)
    n = UltimaFila(ws, 1)
    codigo = CodigoEmpresaA3(codigo)
    For pasada = 1 To 2
        For r = 2 To n
            cod = ValorTexto(ws.Cells(r, 1).Value)
            If (pasada = 1 And cod <> "" And cod <> "*" And CodigoEmpresaA3(cod) = codigo) Or (pasada = 2 And cod = "*") Then
                pct = PorcentajeCelda(ws.Cells(r, 2), ok)
                If ok Then
                    If Not IvaBuscar(pct, cta, dummy, ctaRE) Then
                        pctRE = PorcentajeCelda(ws.Cells(r, 4), ok)
                        If Not ok Then pctRE = 0
                        IvaAgregar pct, ValorTexto(ws.Cells(r, 3).Value), pctRE, ValorTexto(ws.Cells(r, 5).Value)
                    End If
                End If
            End If
        Next r
    Next pasada
End Sub

' Porcentaje de una celda de configuración: si tiene formato % (21 % se guarda
' como 0,21) se multiplica por 100 antes de redondear.
Private Function PorcentajeCelda(ByVal celda As Object, ByRef ok As Boolean) As Currency
    Dim v As Variant
    v = celda.Value
    If VarType(v) = vbDouble And InStr(1, CStr(celda.NumberFormat), "%") > 0 Then
        v = CDbl(v) * 100
    End If
    PorcentajeCelda = LeerImporte(v, "auto", ok)
End Function

' Lee un perfil de la hoja PERFILES (columna con su nombre en la fila 1).
Public Function CargarPerfil(ByVal nombre As String, ByRef per As TPerfil, ByRef msg As String) As Boolean
    Dim ws As Object, col As Long, c As Long, r As Long, n As Long, clave As String, valor As String
    Dim campos(1 To NUM_CAMPOS) As String, i As Long
    per.Nombre = "": per.Campos = "": per.ValoresAbono = "": per.ValoresTicket = ""
    per.ValoresRectificativa = "": per.NombreVacio = "": per.AbonoSiNegativo = False
    If Not ExisteHoja(HOJA_PERFILES) Then
        CargarPerfil = PerfilPredefinido(nombre, per)
        If Not CargarPerfil Then msg = "No existe la hoja " & HOJA_PERFILES & "."
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(HOJA_PERFILES)
    For c = 3 To ws.Cells(1, ws.Columns.Count).End(-4159).Column   ' -4159 = xlToLeft
        If UCase$(ValorTexto(ws.Cells(1, c).Value)) = UCase$(Trim$(nombre)) Then col = c: Exit For
    Next c
    If col = 0 Then
        msg = "El perfil """ & nombre & """ no existe en la hoja " & HOJA_PERFILES & "."
        Exit Function
    End If
    per.Nombre = ValorTexto(ws.Cells(1, col).Value)
    per.Separador = "auto": per.SepDecimal = "auto": per.FormatoFecha = "DMA": per.FilaCabecera = 1
    n = UltimaFila(ws, 1)
    For r = 2 To n
        clave = UCase$(ValorTexto(ws.Cells(r, 1).Value))
        valor = ValorTexto(ws.Cells(r, col).Value)
        For i = 1 To NUM_CAMPOS
            If clave = CampoClave(i) Then campos(i) = valor: Exit For
        Next i
        Select Case clave
            Case "SEPARADOR": If valor <> "" Then per.Separador = valor
            Case "DECIMAL": If valor <> "" Then per.SepDecimal = valor
            Case "FORMATO_FECHA": If valor <> "" Then per.FormatoFecha = UCase$(valor)
            Case "VALORES_ABONO": per.ValoresAbono = valor
            Case "VALORES_TICKET": per.ValoresTicket = valor
            Case "VALORES_RECTIFICATIVA": per.ValoresRectificativa = valor
            Case "NOMBRE_VACIO": per.NombreVacio = valor
            Case "ABONO_SI_NEGATIVO": per.AbonoSiNegativo = EnLista(valor, "SI;S;X;1;VERDADERO;TRUE")
            Case "FILA_CABECERA": If IsNumeric(valor) And valor <> "" Then per.FilaCabecera = CLng(valor)
        End Select
    Next r
    If per.FilaCabecera < 1 Then per.FilaCabecera = 1
    per.Campos = ""
    For i = 1 To NUM_CAMPOS
        If i > 1 Then per.Campos = per.Campos & vbTab
        per.Campos = per.Campos & campos(i)
    Next i
    CargarPerfil = True
End Function

' Nombres de los perfiles definidos (para el desplegable del formulario)
Public Function ListaPerfiles() As Variant
    Dim ws As Object, c As Long, ultima As Long, lista() As String, n As Long
    ReDim lista(0 To 0)
    If Not ExisteHoja(HOJA_PERFILES) Then
        lista(0) = "GENERAL"
        ListaPerfiles = lista
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(HOJA_PERFILES)
    ultima = ws.Cells(1, ws.Columns.Count).End(-4159).Column
    For c = 3 To ultima
        If ValorTexto(ws.Cells(1, c).Value) <> "" Then
            ReDim Preserve lista(0 To n)
            lista(n) = ValorTexto(ws.Cells(1, c).Value)
            n = n + 1
        End If
    Next c
    If n = 0 Then lista(0) = "GENERAL"
    ListaPerfiles = lista
End Function

' =====================================================================
'  CREACIÓN DE LAS HOJAS DE CONFIGURACIÓN (solo si no existen)
' =====================================================================
Public Sub CrearHojasConfiguracion()
    If Not ExisteHoja(HOJA_EMPRESAS) Then CrearHojaEmpresas
    If Not ExisteHoja(HOJA_IVA) Then CrearHojaIVA
    If Not ExisteHoja(HOJA_PERFILES) Then CrearHojaPerfiles
End Sub

Private Function NuevaHoja(ByVal nombre As String) As Object
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nombre
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ActiveWindow.DisplayGridlines = False
    Set NuevaHoja = ws
End Function

Public Sub EstiloCabecera(ByVal rng As Object)
    With rng
        .Font.Bold = True
        .Font.Color = COLOR_BLANCO
        .Interior.Color = COLOR_ACP
        .HorizontalAlignment = -4108          ' xlCenter
        .VerticalAlignment = -4108
        .WrapText = True
    End With
End Sub

Private Sub CrearHojaEmpresas()
    Dim ws As Object, cab As Variant, i As Long, anchos As Variant
    Set ws = NuevaHoja(HOJA_EMPRESAS)
    ws.Tab.Color = COLOR_ACP
    cab = Array("Código empresa a3", "Nombre", "Dígitos plan de cuentas", "Perfil facturas", _
                "Cuenta clientes", "Descripción cuenta clientes", "Cuenta ventas (general)", _
                "Descripción cuenta ventas", "Cuenta ventas si la columna selectora tiene dato", _
                "Descripción cuenta ventas alternativa", "Cuenta retenciones", "Carpeta de salida (última usada)", _
                "Último uso", "Notas")
    anchos = Array(11, 32, 10, 13, 13, 32, 14, 32, 18, 32, 13, 40, 16, 50)
    For i = 0 To UBound(cab)
        ws.Cells(1, i + 1).Value = cab(i)
        ws.Columns(i + 1).ColumnWidth = anchos(i)
    Next i
    EstiloCabecera ws.Range(ws.Cells(1, 1), ws.Cells(1, EMP_NUM_COLUMNAS))
    ws.Rows(1).RowHeight = 45
    ws.Columns(EMP_CODIGO).NumberFormat = "@"
    ws.Range(ws.Columns(EMP_CTA_CLIENTES), ws.Columns(EMP_CTA_RETENCION)).NumberFormat = "@"
    ' Empresa del hotel (flujo del CSV de Contapluz)
    ws.Cells(2, EMP_CODIGO).Value = "01692"
    ws.Cells(2, EMP_NOMBRE).Value = "Patio Posadero"
    ws.Cells(2, EMP_DIGITOS).Value = 8
    ws.Cells(2, EMP_PERFIL).Value = "CONTAPLUZ"
    ws.Cells(2, EMP_CTA_CLIENTES).Value = "43000000"
    ws.Cells(2, EMP_DESC_CLIENTES).Value = "Clientes varios Patio Posadero"
    ws.Cells(2, EMP_CTA_VENTAS).Value = "70500001"
    ws.Cells(2, EMP_DESC_VENTAS).Value = "Prestacion servicios otros"
    ws.Cells(2, EMP_CTA_VENTAS_ALT).Value = "70500000"
    ws.Cells(2, EMP_DESC_VENTAS_ALT).Value = "Prestacion servicios habitacion"
    ws.Cells(2, EMP_CTA_RETENCION).Value = "47300000"
    ws.Cells(2, EMP_NOTAS).Value = "CSV de Contapluz: con habitación (columna Center) -> 70500000; resto -> 70500001"
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub CrearHojaIVA()
    Dim ws As Object, cab As Variant, i As Long, anchos As Variant, datos As Variant
    Set ws = NuevaHoja(HOJA_IVA)
    ws.Tab.Color = COLOR_ACP
    cab = Array("Código empresa a3 (* = todas)", "% IVA", "Cuenta IVA repercutido", "% Recargo equivalencia", _
                "Cuenta recargo", "Notas")
    anchos = Array(14, 8, 16, 12, 16, 50)
    For i = 0 To UBound(cab)
        ws.Cells(1, i + 1).Value = cab(i)
        ws.Columns(i + 1).ColumnWidth = anchos(i)
    Next i
    EstiloCabecera ws.Range(ws.Cells(1, 1), ws.Cells(1, 6))
    ws.Rows(1).RowHeight = 45
    ws.Columns(1).NumberFormat = "@"
    ws.Columns(2).NumberFormat = "0.00"
    ws.Columns(3).NumberFormat = "@"
    ws.Columns(4).NumberFormat = "0.00"
    ws.Columns(5).NumberFormat = "@"
    datos = Array(Array("01692", 4, "47700004"), Array("01692", 10, "47700010"), Array("01692", 21, "47700021"))
    For i = 0 To UBound(datos)
        ws.Cells(i + 2, 1).Value = datos(i)(0)
        ws.Cells(i + 2, 2).Value = datos(i)(1)
        ws.Cells(i + 2, 3).Value = datos(i)(2)
        ws.Cells(i + 2, 4).Value = 0
    Next i
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub CrearHojaPerfiles()
    Dim ws As Object, per As TPerfil, i As Long, nombres As Variant, k As Long, col As Long, r As Long
    Set ws = NuevaHoja(HOJA_PERFILES)
    ws.Tab.Color = COLOR_ACP
    ws.Cells(1, 1).Value = "Campo"
    ws.Cells(1, 2).Value = "Qué es (escribe en cada perfil el nombre de la columna del fichero; alternativas con |; ? delante = opcional)"
    nombres = Array("GENERAL", "CONTAPLUZ")
    For i = 1 To NUM_CAMPOS
        ws.Cells(i + 1, 1).Value = CampoClave(i)
        ws.Cells(i + 1, 2).Value = CampoDescripcion(i)
    Next i
    r = NUM_CAMPOS + 2
    ws.Cells(r, 1).Value = "SEPARADOR": ws.Cells(r, 2).Value = "Separador del CSV: auto, ; , TAB o |"
    ws.Cells(r + 1, 1).Value = "DECIMAL": ws.Cells(r + 1, 2).Value = "Separador decimal de los importes en texto: auto, . o ,"
    ws.Cells(r + 2, 1).Value = "FORMATO_FECHA": ws.Cells(r + 2, 2).Value = "Orden de las fechas en texto: DMA (dd/mm/aaaa), MDA o AMD"
    ws.Cells(r + 3, 1).Value = "VALORES_ABONO": ws.Cells(r + 3, 2).Value = "Valores de la columna TIPO que son abonos (separados por ;)"
    ws.Cells(r + 4, 1).Value = "VALORES_TICKET": ws.Cells(r + 4, 2).Value = "Valores de la columna TIPO que son tickets (separados por ;)"
    ws.Cells(r + 5, 1).Value = "NOMBRE_VACIO": ws.Cells(r + 5, 2).Value = "Nombre de cliente cuando viene vacío"
    ws.Cells(r + 6, 1).Value = "ABONO_SI_NEGATIVO": ws.Cells(r + 6, 2).Value = "SI = sin tipo, un total negativo se trata como abono"
    ws.Cells(r + 7, 1).Value = "FILA_CABECERA": ws.Cells(r + 7, 2).Value = "Fila donde están los nombres de las columnas"
    ws.Cells(r + 8, 1).Value = "VALORES_RECTIFICATIVA": ws.Cells(r + 8, 2).Value = _
        "Valores de TIPO que son rectificativas (tipo 2 con el signo del origen: en negativo disminuye, en positivo aumenta)"
    For k = 0 To UBound(nombres)
        col = 3 + k
        PerfilPredefinido CStr(nombres(k)), per
        ws.Cells(1, col).Value = per.Nombre
        For i = 1 To NUM_CAMPOS
            ws.Cells(i + 1, col).Value = PerfilCampo(per, i)
        Next i
        ws.Cells(r, col).Value = per.Separador
        ws.Cells(r + 1, col).Value = per.SepDecimal
        ws.Cells(r + 2, col).Value = per.FormatoFecha
        ws.Cells(r + 3, col).Value = per.ValoresAbono
        ws.Cells(r + 4, col).Value = per.ValoresTicket
        ws.Cells(r + 5, col).Value = per.NombreVacio
        ws.Cells(r + 6, col).Value = IIf(per.AbonoSiNegativo, "SI", "NO")
        ws.Cells(r + 7, col).Value = per.FilaCabecera
        ws.Cells(r + 8, col).Value = per.ValoresRectificativa
    Next k
    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 60
    ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 22
    ws.Columns(5).ColumnWidth = 22
    ws.Range(ws.Cells(1, 3), ws.Cells(r + 8, 10)).NumberFormat = "@"
    EstiloCabecera ws.Range(ws.Cells(1, 1), ws.Cells(1, 5))
    ws.Range(ws.Cells(2, 1), ws.Cells(r + 8, 1)).Font.Bold = True
    ws.Range(ws.Cells(2, 2), ws.Cells(r + 8, 2)).Font.Color = COLOR_GRIS
    ws.Range(ws.Cells(r, 1), ws.Cells(r + 8, 5)).Interior.Color = COLOR_ACP_CLARO
    ws.Range("C2").Select
    ActiveWindow.FreezePanes = True
End Sub
