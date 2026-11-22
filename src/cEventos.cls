VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' cEventos - FASE 3: eventos de SolidWorks durante la revision.
'   RepaintPostNotify          -> volver a dibujar los tarugos temporales
'   UserSelectionPostNotify    -> sincronizar la seleccion 3D con la lista
'==============================================================================

Private WithEvents mv As SldWorks.ModelView
Private WithEvents ad As SldWorks.AssemblyDoc
Private WithEvents pd As SldWorks.PartDoc            ' pieza multicuerpo (en vez de ensamblaje)

Public Sub Conecta(md As SldWorks.ModelDoc2)
    If md.GetType = swDocPART Then Set pd = md Else Set ad = md
    Set mv = md.GetFirstModelView
End Sub

Public Sub Desconecta()
    Set mv = Nothing
    Set ad = Nothing
    Set pd = Nothing
End Sub

Private Function mv_RepaintPostNotify() As Long
    mPreview.DibujaTodo
    mv_RepaintPostNotify = 0
End Function

Private Function ad_UserSelectionPostNotify() As Long
    mPreview.SincronizaSeleccion
    ad_UserSelectionPostNotify = 0
End Function

Private Function pd_UserSelectionPostNotify() As Long
    mPreview.SincronizaSeleccion
    pd_UserSelectionPostNotify = 0
End Function
