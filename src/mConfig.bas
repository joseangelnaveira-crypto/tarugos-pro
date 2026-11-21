Attribute VB_Name = "mConfig"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mConfig - Parametros de la macro.
' Todos los valores de usuario se guardan en MILIMETROS (lo que ve el usuario).
' La API de SolidWorks trabaja en METROS: usar MmAM() / MAMm() para convertir.
' (Fase 1: valores por defecto. El formulario de configuracion llega en la Fase 8.)
'==============================================================================

Public Const MM As Double = 0.001          ' 1 mm en metros

Public Type TConfig
    Diametro As Double        ' mm
    Prof1 As Double           ' mm  (pieza 1)
    Prof2 As Double           ' mm  (pieza 2)
    DistExtremos As Double    ' mm
    LongRef As Double         ' mm
    NInicial As Long
    IncLong As Double         ' mm
    IncN As Long
    Tol As Double             ' mm  tolerancia geometrica
    AreaMinCara As Double     ' mm2 caras planas menores se ignoran
    MinMaterial As Double     ' mm  material minimo que debe quedar tras el taladro
    SepMinFactor As Double    ' separacion minima entre tarugos = factor x diametro
    TolAngulo As Double       ' coseno minimo para considerar caras paralelas
End Type

Public Cfg As TConfig

Private Const APP_REG As String = "TarugosPro"

' Valores de fabrica; despues se aplican los que el usuario haya guardado en la ventana
' de configuracion (registro de Windows: HKCU\Software\VB and VBA Program Settings).
Public Sub CfgPorDefecto()
    With Cfg
        .Diametro = Leido("Diametro", 8)
        .Prof1 = Leido("Prof1", 31)
        .Prof2 = Leido("Prof2", 11)
        .DistExtremos = Leido("DistExtremos", 50)
        .LongRef = Leido("LongRef", 1000)
        .NInicial = CLng(Leido("NInicial", 3))
        .IncLong = Leido("IncLong", 300)
        .IncN = CLng(Leido("IncN", 1))
        .Tol = Leido("Tol", 0.1)
        .AreaMinCara = 100
        .MinMaterial = 3
        .SepMinFactor = 3
        .TolAngulo = 0.9998
    End With
End Sub

Private Function Leido(ByVal nombre As String, ByVal defecto As Double) As Double
    Dim s As String
    s = GetSetting(APP_REG, "Config", nombre, "")
    If s = "" Then Leido = defecto Else Leido = Val(Replace(s, ",", "."))
End Function

Public Sub CfgGuarda()
    With Cfg
        SaveSetting APP_REG, "Config", "Diametro", Replace(CStr(.Diametro), ",", ".")
        SaveSetting APP_REG, "Config", "Prof1", Replace(CStr(.Prof1), ",", ".")
        SaveSetting APP_REG, "Config", "Prof2", Replace(CStr(.Prof2), ",", ".")
        SaveSetting APP_REG, "Config", "DistExtremos", Replace(CStr(.DistExtremos), ",", ".")
        SaveSetting APP_REG, "Config", "LongRef", Replace(CStr(.LongRef), ",", ".")
        SaveSetting APP_REG, "Config", "NInicial", CStr(.NInicial)
        SaveSetting APP_REG, "Config", "IncLong", Replace(CStr(.IncLong), ",", ".")
        SaveSetting APP_REG, "Config", "IncN", CStr(.IncN)
        SaveSetting APP_REG, "Config", "Tol", Replace(CStr(.Tol), ",", ".")
    End With
End Sub

' Borra lo guardado: vuelven los valores de fabrica
Public Sub CfgRestaura()
    On Error Resume Next
    DeleteSetting APP_REG, "Config"
End Sub

Public Function MmAM(ByVal v As Double) As Double
    MmAM = v * MM
End Function

Public Function MAMm(ByVal v As Double) As Double
    MAMm = v / MM
End Function
