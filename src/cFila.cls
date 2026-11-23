VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFila"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' cFila - una fila de la lista de revision (una casilla por tarugo). Recoge los
' eventos de la casilla y se los pasa al formulario frmRevision.
'==============================================================================

Public Idx As Long
Private WithEvents chk As MSForms.CheckBox

Public Sub Conecta(c As MSForms.CheckBox, ByVal i As Long)
    Set chk = c
    Idx = i
End Sub

Private Sub chk_Change()
    If mPreview.Formulario Is Nothing Then Exit Sub
    mPreview.Formulario.FilaCambia Idx
End Sub

Private Sub chk_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If mPreview.Formulario Is Nothing Then Exit Sub
    mPreview.Formulario.FilaMouseUp Idx, Shift
End Sub

Private Sub chk_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If mPreview.Formulario Is Nothing Then Exit Sub
    mPreview.Formulario.FilaTecla Idx, KeyCode, Shift
End Sub
