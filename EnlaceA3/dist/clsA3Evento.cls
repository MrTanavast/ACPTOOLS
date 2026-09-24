VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsA3Evento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================================
'  ENLACE CONTABLE A3  ·  ACP ASOCIADOS
'  Clase auxiliar del formulario: recoge los eventos de los controles
'  que el formulario crea por código y se los pasa al formulario.
' =====================================================================
Option Explicit

Public WithEvents Boton As MSForms.CommandButton
Attribute Boton.VB_VarHelpID = -1
Public WithEvents Opcion As MSForms.OptionButton
Attribute Opcion.VB_VarHelpID = -1
Public WithEvents Lista As MSForms.ComboBox
Attribute Lista.VB_VarHelpID = -1
Public WithEvents Texto As MSForms.TextBox
Attribute Texto.VB_VarHelpID = -1
Public WithEvents Casilla As MSForms.CheckBox
Attribute Casilla.VB_VarHelpID = -1

Public Nombre As String
Public Formulario As Object

Private Sub Avisar()
    If Not Formulario Is Nothing Then Formulario.ControlEvento Nombre
End Sub

Private Sub Boton_Click()
    Avisar
End Sub

Private Sub Opcion_Click()
    Avisar
End Sub

Private Sub Lista_Change()
    Avisar
End Sub

Private Sub Texto_Change()
    Avisar
End Sub

Private Sub Casilla_Click()
    Avisar
End Sub
