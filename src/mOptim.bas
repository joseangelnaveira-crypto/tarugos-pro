Attribute VB_Name = "mOptim"
Option Explicit
'==============================================================================
' mOptim - FASE 7: progreso, cancelacion y medida de tiempos del analisis.
'
' Usa la barra de progreso nativa de SolidWorks (barra de estado). El usuario
' cancela pulsando ESC; el analisis se detiene sin haber modificado nada.
' Las fases de analisis NO modifican documentos, asi que cancelar es seguro.
'==============================================================================

Public Cancelado As Boolean
Private pb As SldWorks.UserProgressBar
Private maxPb As Long
Private t0 As Double
Public TiempoPiezas As Double
Public TiempoUniones As Double
Public ParesRevisados As Long
Public ParesCandidatos As Long

Public Sub Inicia(app As SldWorks.SldWorks, ByVal texto As String, ByVal maximo As Long)
    Cancelado = False
    maxPb = IIf(maximo < 1, 1, maximo)
    On Error Resume Next
    Set pb = Nothing
    app.GetUserProgressBar pb
    If Not pb Is Nothing Then pb.Start 0, maxPb, texto
    On Error GoTo 0
    t0 = Timer
End Sub

' Devuelve True si el usuario ha cancelado (ESC)
Public Function Paso(ByVal pos As Long) As Boolean
    If pb Is Nothing Then Exit Function
    On Error Resume Next
    Dim r As Long
    r = pb.UpdateProgress(IIf(pos > maxPb, maxPb, pos))
    If r = swUpdateProgressError_UserCancel Then Cancelado = True
    Paso = Cancelado
End Function

Public Function Fin() As Double
    On Error Resume Next
    If Not pb Is Nothing Then pb.End
    Set pb = Nothing
    Fin = Timer - t0
End Function
