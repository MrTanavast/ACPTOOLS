' =====================================================================
'  ENLACE CONTABLE A3  ·  ACP ASOCIADOS
'  CÓDIGO DEL FORMULARIO frmEnlaceA3
'
'  Cómo instalarlo: en el editor de VBA, Insertar > UserForm, cambia su
'  propiedad (Name) a  frmEnlaceA3  y pega TODO este texto en su ventana
'  de código (doble clic en el formulario > borra lo que haya > pegar).
'  No hace falta dibujar nada: los controles se crean al abrirse.
' =====================================================================
Option Explicit

Private mEventos As Collection
Private mLibros() As String
Private mHojas() As String
Private mNHojas As Long
Private mCargando As Boolean
Private mUltimaRuta As String
Private mEmpresaExiste As Boolean
Private mModoAnterior As Integer
Private mVentana As Object

' --- controles ------------------------------------------------------------
Private txtEmpresa As MSForms.TextBox
Private lblEmpresaNombre As MSForms.Label
Private btnConfigurar As MSForms.CommandButton
Private optDiario As MSForms.OptionButton
Private optEmitidas As MSForms.OptionButton
Private optRecibidas As MSForms.OptionButton
Private cmbDigitos As MSForms.ComboBox
Private lblPerfil As MSForms.Label
Private cmbPerfil As MSForms.ComboBox
Private optHoja As MSForms.OptionButton
Private cmbHoja As MSForms.ComboBox
Private optArchivo As MSForms.OptionButton
Private txtArchivo As MSForms.TextBox
Private btnArchivo As MSForms.CommandButton
Private txtCarpeta As MSForms.TextBox
Private btnCarpeta As MSForms.CommandButton
Private chkControl As MSForms.CheckBox
Private chkCopia As MSForms.CheckBox
Private cmbPeriodo As MSForms.ComboBox
Private txtAnio As MSForms.TextBox
Private txtDesde As MSForms.TextBox
Private txtHasta As MSForms.TextBox
Private lblRef As MSForms.Label
Private txtRefDesde As MSForms.TextBox
Private txtRefHasta As MSForms.TextBox
Private lblSeriesIncl As MSForms.Label
Private txtSeriesIncl As MSForms.TextBox
Private lblSeriesExcl As MSForms.Label
Private txtSeriesExcl As MSForms.TextBox
Private lblTipos As MSForms.Label
Private chkFacturas As MSForms.CheckBox
Private chkTickets As MSForms.CheckBox
Private chkAbonos As MSForms.CheckBox
Private chkDescuadrados As MSForms.CheckBox
Private lblAyudaFiltros As MSForms.Label
Private txtResultado As MSForms.TextBox
Private btnAnalizar As MSForms.CommandButton
Private btnGenerar As MSForms.CommandButton
Private btnAbrirCarpeta As MSForms.CommandButton
Private btnCerrar As MSForms.CommandButton

Private Const COL_IZQ_ETQ As Single = 14
Private Const COL_IZQ_CTL As Single = 108
Private Const COL_DER_ETQ As Single = 364
Private Const COL_DER_CTL As Single = 462
Private Const PERIODO_TODO As Long = 0
Private Const PERIODO_ANIO As Long = 1
Private Const PERIODO_PERSONAL As Long = 18

' =====================================================================
'  ARRANQUE
' =====================================================================
Private Sub UserForm_Initialize()
    Dim i As Long, perfiles As Variant, modo As String, emp As String
    mCargando = True
    Set mEventos = New Collection
    On Error Resume Next
    Set mVentana = ActiveWindow
    On Error GoTo 0
    Me.Caption = "Enlace contable a3  ·  ACP Asociados"
    Me.BackColor = COLOR_BLANCO
    Me.Width = 712
    Me.Height = 518
    ConstruirControles

    For i = 6 To 12
        cmbDigitos.AddItem CStr(i)
    Next i
    cmbDigitos.Value = "8"
    perfiles = ListaPerfiles()
    For i = LBound(perfiles) To UBound(perfiles)
        cmbPerfil.AddItem perfiles(i)
    Next i
    If cmbPerfil.ListCount > 0 Then cmbPerfil.ListIndex = 0
    cmbPeriodo.List = Array("Todo el fichero", "Año completo", "1er trimestre", "2º trimestre", "3er trimestre", _
        "4º trimestre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", _
        "Octubre", "Noviembre", "Diciembre", "Personalizado")
    cmbPeriodo.ListIndex = PERIODO_TODO
    txtAnio.Text = CStr(Year(Date))
    chkFacturas.Value = True
    chkTickets.Value = True
    chkAbonos.Value = True
    chkDescuadrados.Value = True
    chkControl.Value = True
    chkCopia.Value = True
    LlenarHojas
    optHoja.Value = True

    modo = LeerPreferencia("A3_ULTIMO_MODO", CStr(MODO_DIARIO))
    If modo = CStr(MODO_EMITIDAS) Then optEmitidas.Value = True Else optDiario.Value = True
    emp = LeerPreferencia("A3_ULTIMA_EMPRESA", "")
    txtEmpresa.Text = emp
    txtResultado.Text = "1. Escribe el código de empresa y elige el tipo de enlace." & vbCrLf & _
                        "2. Elige el origen de los datos (hoja de un libro abierto o un archivo CSV / Excel)." & vbCrLf & _
                        "3. Ajusta los filtros si hace falta y pulsa ANALIZAR para ver el resultado sin escribir nada." & vbCrLf & _
                        "4. Pulsa GENERAR SUENLACE.DAT. Se crea también un Excel de control con las incidencias."
    mCargando = False
    ModoCambiado
    EmpresaCambiada
    AjustarAPantalla
End Sub

' En pantallas pequeñas o con escala grande, reduce la ventana para que quepa
Private Sub AjustarAPantalla()
    Dim z As Double, alto As Double
    On Error Resume Next
    If Application.WindowState = -4137 Then                ' xlMaximized
        alto = Application.UsableHeight + 40
    Else
        alto = Application.Height
    End If
    If alto > 200 And Me.Height > alto - 10 Then
        z = (alto - 10) / Me.Height
        Me.Zoom = Int(z * 100)
        Me.Width = Me.Width * z
        Me.Height = Me.Height * z
    End If
    On Error GoTo 0
End Sub

' Vuelve a poner la ventana delante (al abrir libros, Excel puede taparla)
Private Sub TraerAlFrente()
    On Error Resume Next
    If Not mVentana Is Nothing Then mVentana.Activate
    AppActivate Me.Caption
    On Error GoTo 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    LiberarEventos
End Sub

Private Sub UserForm_Terminate()
    LiberarEventos
End Sub

Private Sub LiberarEventos()
    Dim ev As Object
    If mEventos Is Nothing Then Exit Sub
    For Each ev In mEventos
        Set ev.Formulario = Nothing
    Next ev
    Set mEventos = Nothing
End Sub

' =====================================================================
'  CONSTRUCCIÓN DE LA VENTANA
' =====================================================================
Private Sub ConstruirControles()
    Dim c As Object, ancho As Single

    ancho = 700
    ' --- banda superior ---------------------------------------------------
    Set c = Etiqueta("lblBanda", "", 0, 0, ancho + 20, 48)
    c.BackStyle = 1
    c.BackColor = COLOR_ACP
    Set c = Etiqueta("lblMarca", ChrW(&H2714), 16, 6, 36, 36)
    c.Font.Name = "Segoe UI Symbol"
    c.Font.Size = 24
    c.ForeColor = COLOR_BLANCO
    Set c = Etiqueta("lblTitulo", "Enlace contable a3", 54, 5, 400, 24)
    c.Font.Size = 15
    c.Font.Bold = True
    c.ForeColor = COLOR_BLANCO
    Set c = Etiqueta("lblSubtitulo", "ACP Asociados  ·  Córdoba   —   diarios y libros de facturas  ·  SUENLACE.DAT", 55, 28, 500, 14)
    c.Font.Size = 8.5
    c.ForeColor = COLOR_ACP_MEDIO
    Set c = Etiqueta("lblVersion", "v" & A3_VERSION, ancho - 40, 30, 36, 12)
    c.Font.Size = 7.5
    c.ForeColor = COLOR_ACP_MEDIO

    ' --- 1. empresa y tipo -------------------------------------------------
    Seccion "sec1", "1 · EMPRESA Y TIPO DE ENLACE", COL_IZQ_ETQ, 56, 330
    Etiqueta "lblEmp", "Código empresa", COL_IZQ_ETQ, 78, 90
    Set txtEmpresa = Texto("txtEmpresa", COL_IZQ_CTL, 75, 52)
    txtEmpresa.MaxLength = 5
    txtEmpresa.Font.Bold = True
    txtEmpresa.ControlTipText = "Código de la empresa en a3 (se completa con ceros: 1692 = 01692)"
    Set lblEmpresaNombre = Etiqueta("lblEmpresaNombre", "", COL_IZQ_CTL + 58, 78, 180, 14)
    lblEmpresaNombre.Font.Bold = True
    Set btnConfigurar = Boton("btnConfigurar", "Configurar empresa...", COL_IZQ_CTL + 58, 94, 130, 18)
    btnConfigurar.Font.Size = 8
    btnConfigurar.Visible = False

    Etiqueta "lblTipo", "Tipo de enlace", COL_IZQ_ETQ, 118, 90
    Set optDiario = Opcion("optDiario", "Diario", "modo", COL_IZQ_CTL, 116, 70)
    Set optEmitidas = Opcion("optEmitidas", "Facturas emitidas", "modo", COL_IZQ_CTL + 72, 116, 120)
    Set optRecibidas = Opcion("optRecibidas", "Facturas recibidas (próximamente)", "modo", COL_IZQ_CTL, 136, 200)
    optRecibidas.Enabled = False
    optDiario.ControlTipText = "Asientos sin IVA (plantilla de diario)"
    optEmitidas.ControlTipText = "Libro de facturas de venta: cabecera + detalle de IVA"

    Etiqueta "lblDig", "Dígitos del plan", COL_IZQ_ETQ, 162, 90
    Set cmbDigitos = Lista("cmbDigitos", COL_IZQ_CTL, 159, 46)
    cmbDigitos.ControlTipText = "Longitud de las subcuentas en a3 (570000000 = 9 dígitos)"
    Set lblPerfil = Etiqueta("lblPerfil", "Perfil", COL_IZQ_CTL + 58, 162, 34)
    Set cmbPerfil = Lista("cmbPerfil", COL_IZQ_CTL + 92, 159, 130)
    cmbPerfil.ControlTipText = "Cómo leer el libro de facturas (hoja PERFILES). GENERAL = plantilla de ACP"

    ' --- 2. origen -----------------------------------------------------
    Seccion "sec2", "2 · DATOS DE ORIGEN", COL_IZQ_ETQ, 190, 330
    Set optHoja = Opcion("optHoja", "Hoja abierta", "origen", COL_IZQ_ETQ, 210, 90)
    Set cmbHoja = Lista("cmbHoja", COL_IZQ_CTL, 209, 236)
    cmbHoja.ControlTipText = "Hojas de los libros abiertos en Excel (hoja · libro)"
    cmbHoja.ListWidth = 460
    Set optArchivo = Opcion("optArchivo", "Archivo CSV", "origen", COL_IZQ_ETQ, 234, 90)
    Set txtArchivo = Texto("txtArchivo", COL_IZQ_CTL, 233, 166)
    txtArchivo.ControlTipText = "Ruta del fichero CSV"
    Set btnArchivo = Boton("btnArchivo", "Abrir...", COL_IZQ_CTL + 170, 232, 66, 20)
    btnArchivo.ControlTipText = "Elige un CSV o un Excel (el Excel se abre y aparece en la lista de hojas)"

    ' --- 4. salida ---------------------------------------------------------
    Seccion "sec4", "4 · SALIDA", COL_IZQ_ETQ, 264, 330
    Etiqueta "lblCarpeta", "Carpeta", COL_IZQ_ETQ, 287, 90
    Set txtCarpeta = Texto("txtCarpeta", COL_IZQ_CTL, 284, 166)
    txtCarpeta.ControlTipText = "Carpeta donde se guarda SUENLACE.DAT (se recuerda por empresa)"
    Set btnCarpeta = Boton("btnCarpeta", "Elegir...", COL_IZQ_CTL + 170, 283, 66, 20)
    Set chkControl = Casilla("chkControl", "Excel de control", COL_IZQ_CTL, 308, 110)
    Set chkCopia = Casilla("chkCopia", "Copia del .DAT anterior", COL_IZQ_CTL + 112, 308, 130)
    chkCopia.ControlTipText = "Si ya hay un SUENLACE.DAT en la carpeta, se renombra con fecha y hora"

    ' --- 3. filtros (columna derecha) ---------------------------------------
    Seccion "sec3", "3 · FILTROS", COL_DER_ETQ, 56, 324
    Etiqueta "lblPeriodo", "Periodo", COL_DER_ETQ, 78, 90
    Set cmbPeriodo = Lista("cmbPeriodo", COL_DER_CTL, 75, 118)
    Etiqueta "lblAnio", "Año", COL_DER_CTL + 124, 78, 24
    Set txtAnio = Texto("txtAnio", COL_DER_CTL + 150, 75, 44)
    txtAnio.MaxLength = 4

    Etiqueta "lblDesde", "Fechas desde", COL_DER_ETQ, 102, 90
    Set txtDesde = Texto("txtDesde", COL_DER_CTL, 99, 72)
    Etiqueta "lblHasta", "hasta", COL_DER_CTL + 78, 102, 30
    Set txtHasta = Texto("txtHasta", COL_DER_CTL + 110, 99, 72)
    txtDesde.ControlTipText = "dd/mm/aaaa (vacío = sin límite)"
    txtHasta.ControlTipText = "dd/mm/aaaa (vacío = sin límite)"

    Set lblRef = Etiqueta("lblRef", "Documentos desde", COL_DER_ETQ, 126, 96)
    Set txtRefDesde = Texto("txtRefDesde", COL_DER_CTL, 123, 72)
    Etiqueta "lblRefHasta", "hasta", COL_DER_CTL + 78, 126, 30
    Set txtRefHasta = Texto("txtRefHasta", COL_DER_CTL + 110, 123, 72)

    Set lblSeriesIncl = Etiqueta("lblSeriesIncl", "Solo las series", COL_DER_ETQ, 150, 96)
    Set txtSeriesIncl = Texto("txtSeriesIncl", COL_DER_CTL, 147, 182)
    txtSeriesIncl.ControlTipText = "Series separadas por ; (p.ej. F;R o INV;REC). Vacío = todas"
    Set lblSeriesExcl = Etiqueta("lblSeriesExcl", "Excluir series", COL_DER_ETQ, 174, 96)
    Set txtSeriesExcl = Texto("txtSeriesExcl", COL_DER_CTL, 171, 182)
    txtSeriesExcl.ControlTipText = "Series que NO se exportan, separadas por ;"
    Set chkDescuadrados = Casilla("chkDescuadrados", "Excluir asientos descuadrados (recomendado)", COL_DER_CTL, 148, 220)

    Set lblTipos = Etiqueta("lblTipos", "Incluir", COL_DER_ETQ, 198, 90)
    Set chkFacturas = Casilla("chkFacturas", "Facturas", COL_DER_CTL, 196, 64)
    Set chkTickets = Casilla("chkTickets", "Tickets", COL_DER_CTL + 66, 196, 58)
    Set chkAbonos = Casilla("chkAbonos", "Abonos", COL_DER_CTL + 126, 196, 60)

    Set lblAyudaFiltros = Etiqueta("lblAyudaFiltros", "", COL_DER_ETQ, 222, 324, 60)
    lblAyudaFiltros.Font.Size = 8
    lblAyudaFiltros.ForeColor = COLOR_GRIS
    lblAyudaFiltros.WordWrap = True

    ' --- resultado y botones ----------------------------------------------
    Set c = Etiqueta("lblResultado", "RESULTADO", COL_IZQ_ETQ, 330, 120)
    c.Font.Bold = True
    c.ForeColor = COLOR_ACP
    Set txtResultado = Texto("txtResultado", COL_IZQ_ETQ, 346, ancho - 28)
    txtResultado.Height = 98
    txtResultado.MultiLine = True
    txtResultado.WordWrap = True
    txtResultado.ScrollBars = 2
    txtResultado.Locked = True
    txtResultado.BackColor = COLOR_ACP_CLARO
    txtResultado.Font.Name = "Consolas"
    txtResultado.Font.Size = 8.5
    txtResultado.SpecialEffect = 0
    txtResultado.BorderStyle = 1
    txtResultado.BorderColor = COLOR_ACP_MEDIO

    Set btnAnalizar = Boton("btnAnalizar", "Analizar", COL_IZQ_ETQ, 452, 110, 28)
    btnAnalizar.BackColor = COLOR_ACP_CLARO
    btnAnalizar.ForeColor = COLOR_ACP
    btnAnalizar.Font.Bold = True
    btnAnalizar.ControlTipText = "Lee y comprueba los datos sin escribir ningún fichero"
    Set btnGenerar = Boton("btnGenerar", "GENERAR SUENLACE.DAT", COL_IZQ_ETQ + 118, 452, 200, 28)
    btnGenerar.BackColor = COLOR_ACP
    btnGenerar.ForeColor = COLOR_BLANCO
    btnGenerar.Font.Bold = True
    btnGenerar.Font.Size = 10
    Set btnAbrirCarpeta = Boton("btnAbrirCarpeta", "Abrir carpeta", ancho - 212, 452, 100, 28)
    btnAbrirCarpeta.Enabled = False
    Set btnCerrar = Boton("btnCerrar", "Cerrar", ancho - 106, 452, 92, 28)
    btnCerrar.Cancel = True
End Sub

' --- ayudantes de construcción ----------------------------------------------
Private Function Nuevo(ByVal progId As String, ByVal nombre As String, ByVal x As Single, ByVal y As Single, _
                       ByVal ancho As Single, ByVal alto As Single) As Object
    Dim c As Object
    Set c = Me.Controls.Add(progId, nombre, True)
    c.Left = x
    c.Top = y
    c.Width = ancho
    c.Height = alto
    On Error Resume Next
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    On Error GoTo 0
    Set Nuevo = c
End Function

Private Function Etiqueta(ByVal nombre As String, ByVal texto As String, ByVal x As Single, ByVal y As Single, _
                          ByVal ancho As Single, Optional ByVal alto As Single = 14) As MSForms.Label
    Dim c As MSForms.Label
    Set c = Nuevo("Forms.Label.1", nombre, x, y, ancho, alto)
    c.Caption = texto
    c.BackStyle = 0
    c.ForeColor = COLOR_TEXTO
    c.WordWrap = False
    Set Etiqueta = c
End Function

Private Sub Seccion(ByVal nombre As String, ByVal titulo As String, ByVal x As Single, ByVal y As Single, ByVal ancho As Single)
    Dim c As MSForms.Label
    Set c = Etiqueta(nombre, titulo, x, y, ancho, 14)
    c.Font.Bold = True
    c.ForeColor = COLOR_ACP
    Set c = Etiqueta(nombre & "_linea", "", x, y + 14, ancho, 1)
    c.BackStyle = 1
    c.BackColor = COLOR_ACP_MEDIO
End Sub

Private Function Texto(ByVal nombre As String, ByVal x As Single, ByVal y As Single, ByVal ancho As Single) As MSForms.TextBox
    Dim c As MSForms.TextBox
    Set c = Nuevo("Forms.TextBox.1", nombre, x, y, ancho, 18)
    c.SpecialEffect = 0
    c.BorderStyle = 1
    c.BorderColor = COLOR_ACP_MEDIO
    Enganchar c, nombre
    Set Texto = c
End Function

Private Function Lista(ByVal nombre As String, ByVal x As Single, ByVal y As Single, ByVal ancho As Single) As MSForms.ComboBox
    Dim c As MSForms.ComboBox
    Set c = Nuevo("Forms.ComboBox.1", nombre, x, y, ancho, 18)
    c.Style = 2                                  ' fmStyleDropDownList
    c.SpecialEffect = 0
    c.BorderStyle = 1
    c.BorderColor = COLOR_ACP_MEDIO
    c.ListRows = 20
    Enganchar c, nombre
    Set Lista = c
End Function

Private Function Opcion(ByVal nombre As String, ByVal texto As String, ByVal grupo As String, ByVal x As Single, _
                        ByVal y As Single, ByVal ancho As Single) As MSForms.OptionButton
    Dim c As MSForms.OptionButton
    Set c = Nuevo("Forms.OptionButton.1", nombre, x, y, ancho, 18)
    c.Caption = texto
    c.GroupName = grupo
    c.BackStyle = 0
    c.ForeColor = COLOR_TEXTO
    Enganchar c, nombre
    Set Opcion = c
End Function

Private Function Casilla(ByVal nombre As String, ByVal texto As String, ByVal x As Single, ByVal y As Single, _
                         ByVal ancho As Single) As MSForms.CheckBox
    Dim c As MSForms.CheckBox
    Set c = Nuevo("Forms.CheckBox.1", nombre, x, y, ancho, 18)
    c.Caption = texto
    c.BackStyle = 0
    c.ForeColor = COLOR_TEXTO
    Enganchar c, nombre
    Set Casilla = c
End Function

Private Function Boton(ByVal nombre As String, ByVal texto As String, ByVal x As Single, ByVal y As Single, _
                       ByVal ancho As Single, ByVal alto As Single) As MSForms.CommandButton
    Dim c As MSForms.CommandButton
    Set c = Nuevo("Forms.CommandButton.1", nombre, x, y, ancho, alto)
    c.Caption = texto
    c.TakeFocusOnClick = False
    Enganchar c, nombre
    Set Boton = c
End Function

Private Sub Enganchar(ByVal ctl As Object, ByVal nombre As String)
    Dim ev As clsA3Evento
    Set ev = New clsA3Evento
    ev.Nombre = nombre
    Set ev.Formulario = Me
    Select Case TypeName(ctl)
        Case "CommandButton": Set ev.Boton = ctl
        Case "OptionButton": Set ev.Opcion = ctl
        Case "ComboBox": Set ev.Lista = ctl
        Case "TextBox": Set ev.Texto = ctl
        Case "CheckBox": Set ev.Casilla = ctl
    End Select
    mEventos.Add ev
End Sub

' =====================================================================
'  EVENTOS (los llama clsA3Evento)
' =====================================================================
Public Sub ControlEvento(ByVal nombre As String)
    If mCargando Then Exit Sub
    On Error GoTo Fallo
    Select Case nombre
        Case "txtEmpresa"
            EmpresaCambiada
        Case "optDiario", "optEmitidas"
            ModoCambiado
            btnConfigurar.Visible = (ModoActual() = MODO_EMITIDAS) And Not mEmpresaExiste And _
                                    SoloDigitos(Trim$(txtEmpresa.Text))
        Case "cmbPeriodo", "txtAnio"
            PeriodoCambiado
        Case "txtDesde", "txtHasta"
            mCargando = True
            cmbPeriodo.ListIndex = PERIODO_PERSONAL
            mCargando = False
        Case "cmbHoja"
            mCargando = True
            optHoja.Value = True
            mCargando = False
            CarpetaPorDefecto
        Case "txtArchivo"
            mCargando = True
            If Trim$(txtArchivo.Text) <> "" Then optArchivo.Value = True
            mCargando = False
        Case "btnArchivo"
            ElegirArchivo
        Case "btnCarpeta"
            ElegirCarpeta
        Case "btnConfigurar"
            ConfigurarEmpresa
        Case "btnAnalizar"
            Ejecutar True
        Case "btnGenerar"
            Ejecutar False
        Case "btnAbrirCarpeta"
            AbrirCarpeta
        Case "btnCerrar"
            Unload Me
    End Select
    Exit Sub
Fallo:
    mCargando = False
    MsgBox "Error: " & Err.Description, vbExclamation, "Enlace contable a3"
End Sub

Private Function PerfilElegido() As String
    If cmbPerfil.ListIndex >= 0 Then
        PerfilElegido = cmbPerfil.List(cmbPerfil.ListIndex)
    Else
        PerfilElegido = "GENERAL"
    End If
End Function

Private Function ModoActual() As Integer
    If optEmitidas.Value Then ModoActual = MODO_EMITIDAS Else ModoActual = MODO_DIARIO
End Function

Private Sub ModoCambiado()
    Dim fact As Boolean
    fact = (ModoActual() = MODO_EMITIDAS)
    ' el rango de documentos / asientos cambia de significado: se vacía
    If mModoAnterior <> 0 And mModoAnterior <> ModoActual() Then
        mCargando = True
        txtRefDesde.Text = ""
        txtRefHasta.Text = ""
        mCargando = False
    End If
    mModoAnterior = ModoActual()
    lblPerfil.Visible = fact
    cmbPerfil.Visible = fact
    lblSeriesIncl.Visible = fact
    txtSeriesIncl.Visible = fact
    lblSeriesExcl.Visible = fact
    txtSeriesExcl.Visible = fact
    lblTipos.Visible = fact
    chkFacturas.Visible = fact
    chkTickets.Visible = fact
    chkAbonos.Visible = fact
    chkDescuadrados.Visible = Not fact
    If fact Then
        lblRef.Caption = "Documentos desde"
        txtRefDesde.ControlTipText = "Primer documento, p.ej. INV1510 (misma serie que el último)"
        txtRefHasta.ControlTipText = "Último documento, p.ej. INV1704"
        lblAyudaFiltros.Caption = "Deja vacío lo que no quieras filtrar. El rango de documentos compara la serie " & _
            "y el número (INV1510 a INV1704). Los anulados y los importes a 0 se excluyen siempre."
    Else
        lblRef.Caption = "Asientos desde"
        txtRefDesde.ControlTipText = "Primer nº de asiento a exportar"
        txtRefHasta.ControlTipText = "Último nº de asiento a exportar"
        lblAyudaFiltros.Caption = "Deja vacío lo que no quieras filtrar. Un asiento con errores o descuadrado " & _
            "se excluye entero y aparece en el Excel de control."
    End If
End Sub

Private Sub EmpresaCambiada()
    Dim emp As TEmpresa, cod As String, i As Long
    cod = Trim$(txtEmpresa.Text)
    btnConfigurar.Visible = False
    mEmpresaExiste = False
    If cod = "" Then
        lblEmpresaNombre.Caption = ""
    ElseIf Not SoloDigitos(cod) Then
        lblEmpresaNombre.Caption = "Solo números"
        lblEmpresaNombre.ForeColor = COLOR_ERROR
    Else
        mEmpresaExiste = CargarEmpresa(cod, emp)
        mCargando = True
        If mEmpresaExiste Then
            lblEmpresaNombre.Caption = IIf(emp.Nombre <> "", emp.Nombre, "(sin nombre)")
            lblEmpresaNombre.ForeColor = COLOR_OK
            If emp.Digitos >= 6 And emp.Digitos <= 12 Then cmbDigitos.Value = CStr(emp.Digitos)
        Else
            lblEmpresaNombre.Caption = "Empresa " & CodigoEmpresaA3(cod) & " no configurada"
            lblEmpresaNombre.ForeColor = COLOR_AVISO
            btnConfigurar.Visible = (ModoActual() = MODO_EMITIDAS)
        End If
        ' perfil y carpeta SIEMPRE de la empresa escrita (nunca los de la anterior)
        For i = 0 To cmbPerfil.ListCount - 1
            If UCase$(cmbPerfil.List(i)) = UCase$(emp.Perfil) Then cmbPerfil.ListIndex = i
        Next i
        txtCarpeta.Text = emp.CarpetaSalida
        mCargando = False
        If emp.CarpetaSalida = "" Then CarpetaPorDefecto
    End If
    lblEmpresaNombre.ControlTipText = lblEmpresaNombre.Caption
End Sub

Private Sub PeriodoCambiado()
    Dim anio As Long, k As Long, d1 As Date, d2 As Date
    k = cmbPeriodo.ListIndex
    If k = PERIODO_PERSONAL Or k < 0 Then Exit Sub
    anio = CLng(Val(txtAnio.Text))
    If anio < 1990 Or anio > 2099 Then anio = Year(Date)
    mCargando = True
    Select Case k
        Case PERIODO_TODO
            txtDesde.Text = ""
            txtHasta.Text = ""
        Case PERIODO_ANIO
            d1 = DateSerial(anio, 1, 1): d2 = DateSerial(anio, 12, 31)
        Case 2 To 5
            d1 = DateSerial(anio, (k - 2) * 3 + 1, 1): d2 = DateSerial(anio, (k - 2) * 3 + 4, 0)
        Case 6 To 17
            d1 = DateSerial(anio, k - 5, 1): d2 = DateSerial(anio, k - 4, 0)
    End Select
    If k <> PERIODO_TODO Then
        txtDesde.Text = FechaTexto(d1)
        txtHasta.Text = FechaTexto(d2)
    End If
    mCargando = False
End Sub

' Lista "[libro] hoja" de todos los libros abiertos visibles
Private Sub LlenarHojas()
    Dim wb As Object, ws As Object, visible As Boolean, seleccion As Long, actual As String, libro As String, activa As String, i As Long
    mNHojas = 0
    ReDim mLibros(1 To 16)
    ReDim mHojas(1 To 16)
    mCargando = True
    cmbHoja.Clear
    On Error Resume Next
    actual = ActiveWorkbook.Name & "|" & ActiveSheet.Name
    On Error GoTo 0
    For Each wb In Application.Workbooks
        visible = True
        On Error Resume Next
        visible = wb.Windows(1).Visible
        On Error GoTo 0
        If visible Then
            For Each ws In wb.Worksheets
                If Not (wb.Name = ThisWorkbook.Name And EsHojaConfiguracion(ws.Name)) Then
                    mNHojas = mNHojas + 1
                    If mNHojas > UBound(mLibros) Then
                        ReDim Preserve mLibros(1 To mNHojas * 2)
                        ReDim Preserve mHojas(1 To mNHojas * 2)
                    End If
                    mLibros(mNHojas) = wb.Name
                    mHojas(mNHojas) = ws.Name
                    cmbHoja.AddItem ws.Name & "   ·   " & wb.Name
                    If wb.Name & "|" & ws.Name = actual Then seleccion = mNHojas
                End If
            Next ws
        End If
    Next wb
    ' Si la hoja activa es de la herramienta, se propone la hoja activa del último
    ' libro de datos abierto (nunca las plantillas de ejemplo de la herramienta)
    If seleccion = 0 Then
        For i = mNHojas To 1 Step -1
            If mLibros(i) <> ThisWorkbook.Name Then
                libro = mLibros(i)
                Exit For
            End If
        Next i
        If libro <> "" Then
            On Error Resume Next
            activa = Application.Workbooks(libro).ActiveSheet.Name
            On Error GoTo 0
            For i = 1 To mNHojas
                If mLibros(i) = libro Then
                    If seleccion = 0 Or mHojas(i) = activa Then seleccion = i
                End If
            Next i
        End If
    End If
    If seleccion > 0 Then cmbHoja.ListIndex = seleccion - 1
    mCargando = False
End Sub

Private Function EsHojaConfiguracion(ByVal nombre As String) As Boolean
    Select Case UCase$(nombre)
        Case HOJA_INICIO, HOJA_EMPRESAS, HOJA_IVA, HOJA_PERFILES
            EsHojaConfiguracion = True
    End Select
End Function

Private Sub SeleccionarHoja(ByVal libro As String, ByVal hoja As String)
    Dim i As Long
    For i = 1 To mNHojas
        If mLibros(i) = libro And (hoja = "" Or mHojas(i) = hoja) Then
            mCargando = True
            cmbHoja.ListIndex = i - 1
            optHoja.Value = True
            mCargando = False
            Exit Sub
        End If
    Next i
End Sub

Private Sub ElegirArchivo()
    Dim ruta As Variant, ext As String, wb As Object, nombre As String
    ruta = Application.GetOpenFilename( _
        "Datos (*.csv;*.txt;*.xlsx;*.xlsm;*.xls),*.csv;*.txt;*.xlsx;*.xlsm;*.xls,Todos (*.*),*.*", , _
        "Elige el diario o el libro de facturas")
    If VarType(ruta) = vbBoolean Then Exit Sub
    ext = LCase$(Mid$(ruta, InStrRev(ruta, ".") + 1))
    If ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "xlsb" Then
        nombre = NombreDeRuta(CStr(ruta))
        On Error Resume Next
        Set wb = Application.Workbooks(nombre)
        On Error GoTo 0
        If Not wb Is Nothing Then
            If StrComp(wb.FullName, CStr(ruta), vbTextCompare) <> 0 Then
                If LCase$(Left$(wb.FullName, 4)) = "http" Then
                    ' OneDrive / SharePoint: la ruta es una URL y no se puede comparar
                    If MsgBox("Ya está abierto un libro llamado """ & nombre & """:" & vbCrLf & wb.FullName & vbCrLf & vbCrLf & _
                              "¿Es el mismo archivo?", vbQuestion + vbYesNo, "Enlace contable a3") = vbNo Then Exit Sub
                Else
                    MsgBox "Ya hay abierto OTRO libro llamado """ & nombre & """:" & vbCrLf & wb.FullName & vbCrLf & vbCrLf & _
                           "Excel no puede abrir dos libros con el mismo nombre. Ciérralo y vuelve a elegir el archivo.", _
                           vbExclamation, "Enlace contable a3"
                    Exit Sub
                End If
            End If
        Else
            Set wb = Application.Workbooks.Open(Filename:=CStr(ruta), UpdateLinks:=0, ReadOnly:=True)
            TraerAlFrente
        End If
        LlenarHojas
        SeleccionarHoja wb.Name, wb.ActiveSheet.Name
    Else
        mCargando = True
        txtArchivo.Text = CStr(ruta)
        optArchivo.Value = True
        mCargando = False
    End If
    mUltimaRuta = CarpetaDeRuta(CStr(ruta))
    If Trim$(txtCarpeta.Text) = "" Then txtCarpeta.Text = mUltimaRuta
End Sub

Private Sub CarpetaPorDefecto()
    Dim i As Long, wb As Object
    If Trim$(txtCarpeta.Text) <> "" Then Exit Sub
    If optArchivo.Value And Trim$(txtArchivo.Text) <> "" Then
        txtCarpeta.Text = CarpetaDeRuta(Trim$(txtArchivo.Text))
        Exit Sub
    End If
    i = cmbHoja.ListIndex + 1
    If i < 1 Or i > mNHojas Then Exit Sub
    On Error Resume Next
    Set wb = Application.Workbooks(mLibros(i))
    If Not wb Is Nothing Then
        If wb.Path <> "" Then txtCarpeta.Text = wb.Path
    End If
    On Error GoTo 0
End Sub

Private Sub ElegirCarpeta()
    Dim fd As Object
    Set fd = Application.FileDialog(4)                    ' msoFileDialogFolderPicker
    fd.Title = "Carpeta donde se guardará SUENLACE.DAT"
    If Trim$(txtCarpeta.Text) <> "" Then fd.InitialFileName = Trim$(txtCarpeta.Text) & "\"
    If fd.Show = -1 Then txtCarpeta.Text = fd.SelectedItems(1)
End Sub

Private Sub AbrirCarpeta()
    Dim carpeta As String
    carpeta = Trim$(txtCarpeta.Text)
    If carpeta = "" Then Exit Sub
    Shell "explorer.exe """ & carpeta & """", vbNormalFocus
End Sub

Private Sub ConfigurarEmpresa()
    Dim cod As String, nombre As String, emp As TEmpresa
    cod = Trim$(txtEmpresa.Text)
    If cod = "" Or Not SoloDigitos(cod) Then
        MsgBox "Escribe primero el código de la empresa (solo números).", vbExclamation, "Enlace contable a3"
        Exit Sub
    End If
    nombre = InputBox("Nombre de la empresa " & CodigoEmpresaA3(cod) & ":", "Alta de empresa")
    If Trim$(nombre) = "" Then Exit Sub
    AltaEmpresaPorDefecto cod, Trim$(nombre), CInt(Val(cmbDigitos.Value)), PerfilElegido()
    EmpresaCambiada
    CargarEmpresa cod, emp
    If MsgBox("Empresa dada de alta con cuentas por defecto (fila marcada en naranja en EMPRESAS y en IVA):" & vbCrLf & vbCrLf & _
              "  Clientes " & emp.CtaClientes & "   Ventas " & emp.CtaVentas & "   IVA 477... por tipo" & vbCrLf & vbCrLf & _
              "Revisa las cuentas antes de generar el fichero." & vbCrLf & _
              "¿Quieres revisarlas ahora? (se cerrará esta ventana)", vbQuestion + vbYesNo, "Enlace contable a3") = vbYes Then
        Unload Me
        ThisWorkbook.Worksheets(HOJA_EMPRESAS).Activate
        ThisWorkbook.Worksheets(HOJA_EMPRESAS).Cells(emp.Fila, 1).Select
    End If
End Sub

' =====================================================================
'  ANALIZAR / GENERAR
' =====================================================================
Private Sub Ejecutar(ByVal soloAnalizar As Boolean)
    Dim p As TParametros, fil As TFiltros, informe As String, ok As Boolean, i As Long, rutaDat As String

    ' --- parámetros --------------------------------------------------------
    p.Empresa = Trim$(txtEmpresa.Text)
    If p.Empresa = "" Or Not SoloDigitos(p.Empresa) Then
        MsgBox "Escribe el código de empresa de a3 (solo números).", vbExclamation, "Enlace contable a3"
        txtEmpresa.SetFocus
        Exit Sub
    End If
    p.Digitos = CInt(Val(cmbDigitos.Value))
    p.Modo = ModoActual()
    p.Perfil = PerfilElegido()
    If optArchivo.Value Then
        p.OrigenTipo = ORIGEN_CSV
        p.OrigenRuta = Trim$(txtArchivo.Text)
        If p.OrigenRuta = "" Then
            MsgBox "Elige el archivo de origen.", vbExclamation, "Enlace contable a3"
            Exit Sub
        End If
    Else
        p.OrigenTipo = ORIGEN_HOJA
        i = cmbHoja.ListIndex + 1
        If i < 1 Or i > mNHojas Then
            MsgBox "Elige la hoja con los datos.", vbExclamation, "Enlace contable a3"
            Exit Sub
        End If
        p.OrigenLibro = mLibros(i)
        p.OrigenHoja = mHojas(i)
    End If
    p.CarpetaSalida = Trim$(txtCarpeta.Text)
    If Right$(p.CarpetaSalida, 1) = "\" And Len(p.CarpetaSalida) > 3 Then p.CarpetaSalida = Left$(p.CarpetaSalida, Len(p.CarpetaSalida) - 1)
    If Not soloAnalizar And p.CarpetaSalida = "" Then
        MsgBox "Indica la carpeta donde se guardará SUENLACE.DAT.", vbExclamation, "Enlace contable a3"
        Exit Sub
    End If
    p.GenerarControl = chkControl.Value
    p.CopiaSeguridad = chkCopia.Value

    ' --- filtros ------------------------------------------------------------
    If Not LeerFechaFiltro(txtDesde.Text, fil.UsarDesde, fil.FechaDesde, "Fecha desde") Then Exit Sub
    If Not LeerFechaFiltro(txtHasta.Text, fil.UsarHasta, fil.FechaHasta, "Fecha hasta") Then Exit Sub
    If fil.UsarDesde And fil.UsarHasta Then
        If fil.FechaDesde > fil.FechaHasta Then
            MsgBox "La fecha desde es posterior a la fecha hasta.", vbExclamation, "Enlace contable a3"
            Exit Sub
        End If
    End If
    fil.RefDesde = Trim$(txtRefDesde.Text)
    fil.RefHasta = Trim$(txtRefHasta.Text)
    If p.Modo = MODO_EMITIDAS Then
        fil.SeriesIncluir = Trim$(txtSeriesIncl.Text)
        fil.SeriesExcluir = Trim$(txtSeriesExcl.Text)
        fil.InclFacturas = chkFacturas.Value
        fil.InclTickets = chkTickets.Value
        fil.InclAbonos = chkAbonos.Value
        If Not (fil.InclFacturas Or fil.InclTickets Or fil.InclAbonos) Then
            MsgBox "Marca al menos un tipo de documento (facturas, tickets o abonos).", vbExclamation, "Enlace contable a3"
            Exit Sub
        End If
    Else
        fil.InclFacturas = True
        fil.InclTickets = True
        fil.InclAbonos = True
        fil.ExcluirDescuadrados = chkDescuadrados.Value
    End If

    ' --- ejecución ---------------------------------------------------------
    txtResultado.Text = IIf(soloAnalizar, "Analizando...", "Generando...")
    txtResultado.ForeColor = COLOR_TEXTO
    DoEvents
    ok = EjecutarEnlace(p, fil, soloAnalizar, informe, rutaDat)
    If Not soloAnalizar Then TraerAlFrente
    txtResultado.Text = informe
    txtResultado.SelStart = 0
    If Not ok Then
        txtResultado.ForeColor = COLOR_ERROR
    ElseIf IncContar(INC_EXCLUIDO) > 0 Then
        txtResultado.ForeColor = COLOR_AVISO
    Else
        txtResultado.ForeColor = COLOR_OK
    End If
    If rutaDat <> "" Then
        btnAbrirCarpeta.Enabled = True
        MsgBox "SUENLACE.DAT generado en:" & vbCrLf & rutaDat & vbCrLf & vbCrLf & _
               IIf(p.GenerarControl, "El Excel de control está abierto detrás de esta ventana." & vbCrLf, "") & _
               "Recuerda: en a3, chequea (CH) antes de enlazar.", vbInformation, "Enlace contable a3"
    End If
End Sub

Private Function LeerFechaFiltro(ByVal texto As String, ByRef usar As Boolean, ByRef fecha As Date, ByVal nombre As String) As Boolean
    Dim ok As Boolean, d As Date
    texto = Trim$(texto)
    usar = False
    If texto = "" Then
        LeerFechaFiltro = True
        Exit Function
    End If
    d = ParseFechaTexto(texto, "DMA", ok)
    If Not ok Then
        MsgBox nombre & ": """ & texto & """ no es una fecha válida (dd/mm/aaaa).", vbExclamation, "Enlace contable a3"
        Exit Function
    End If
    usar = True
    fecha = d
    LeerFechaFiltro = True
End Function
