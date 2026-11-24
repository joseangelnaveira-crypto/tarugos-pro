Attribute VB_Name = "mConfig"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mConfig - Parametros de la macro.
' Todos los valores de usuario se guardan en MILIMETROS (lo que ve el usuario).
' La API de SolidWorks trabaja en METROS: multiplicar por MM (0,001) para convertir de mm.
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
    Tol As Double             ' mm  holgura maxima entre piezas para considerarlas en contacto
    AreaMinCara As Double     ' mm2 caras planas menores se ignoran
    MinMaterial As Double     ' mm  material minimo que debe quedar tras el taladro
    SepMinFactor As Double    ' separacion minima entre tarugos = factor x diametro
    TolAngulo As Double       ' coseno minimo para considerar caras paralelas
End Type

Public Cfg As TConfig

Private Const APP_REG As String = "TarugosPro"

' Valores de fabrica; despues se aplican los que el usuario haya guardado en la ventana
' de configuracion (registro de Windows: HKCU\Software\VB and VBA Program Settings).
Private ignoraRegistro As Boolean

Public Sub CfgPorDefecto()
    ' Se leen los ajustes guardados; si alguno esta fuera de rango (archivo del registro
    ' modificado o corrupto) se descartan todos y se usan los valores de fabrica.
    ignoraRegistro = False
    CargaCfg
    With Cfg
        If ValidaParametros(.Diametro, .Prof1, .Prof2, .DistExtremos, .LongRef, .NInicial, .IncLong, .IncN, .Tol) <> "" Then
            ignoraRegistro = True
            CargaCfg
            ignoraRegistro = False
        End If
    End With
End Sub

' Comprueba los parametros de usuario. Devuelve "" si todos son validos o el motivo del rechazo.
' (limites razonables para tarugos de mueble; evitan valores que producirian cortes absurdos)
Public Function ValidaParametros(ByVal d As Double, ByVal p1 As Double, ByVal p2 As Double, ByVal de As Double, _
                                 ByVal lr As Double, ByVal ni As Double, ByVal il As Double, ByVal inn As Double, _
                                 ByVal tl As Double) As String
    If d < 1 Or d > 30 Then ValidaParametros = "El diametro debe estar entre 1 y 30 mm.": Exit Function
    If p1 < 1 Or p1 > 200 Then ValidaParametros = "La profundidad de la pieza 1 debe estar entre 1 y 200 mm.": Exit Function
    If p2 < 1 Or p2 > 200 Then ValidaParametros = "La profundidad de la pieza 2 debe estar entre 1 y 200 mm.": Exit Function
    If de < 0 Or de > 1000 Then ValidaParametros = "La distancia a los extremos debe estar entre 0 y 1000 mm.": Exit Function
    If lr < 1 Or lr > 20000 Then ValidaParametros = "La longitud de referencia debe estar entre 1 y 20000 mm.": Exit Function
    If ni < 1 Or ni > 20 Or ni <> Int(ni) Then ValidaParametros = "El numero inicial de tarugos debe ser un entero entre 1 y 20.": Exit Function
    If il < 1 Or il > 20000 Then ValidaParametros = "El incremento de longitud debe estar entre 1 y 20000 mm.": Exit Function
    If inn < 0 Or inn > 10 Or inn <> Int(inn) Then ValidaParametros = "El incremento de tarugos debe ser un entero entre 0 y 10.": Exit Function
    If tl < 0.001 Or tl > 5 Then ValidaParametros = "La holgura entre piezas debe estar entre 0,001 y 5 mm.": Exit Function
End Function

Private Sub CargaCfg()
    With Cfg
        .Diametro = Leido("Diametro", 8)
        .Prof1 = Leido("Prof1", 31)
        .Prof2 = Leido("Prof2", 11)
        .DistExtremos = Leido("DistExtremos", 50)
        .LongRef = Leido("LongRef", 1000)
        .NInicial = LeidoL("NInicial", 3)
        .IncLong = Leido("IncLong", 300)
        .IncN = LeidoL("IncN", 1)
        .Tol = Leido("Tol", 0.2)
        .AreaMinCara = 100
        .MinMaterial = 3
        ' 6,25 x diametro = 50 mm con tarugos de 8: antes 3 (24 mm) y en uniones cortas salian tarugos
        ' pegados (3 en 150 mm, a 25 mm). Con 50 mm una union de 150 mm lleva 2 tarugos (no se queda
        ' en 1, que dejaria girar la pieza) y las uniones de 200 mm o mas siguen con 3.
        .SepMinFactor = 6.25
        .TolAngulo = 0.9998
    End With
End Sub

Private Function Leido(ByVal nombre As String, ByVal defecto As Double) As Double
    Dim s As String
    s = GetSetting(APP_REG, "Config", nombre, "")
    If ignoraRegistro Or s = "" Then Leido = defecto Else Leido = Val(Replace(s, ",", "."))
End Function

' Entero leido del registro; un valor fuera del rango de Long no provoca error (se usa el defecto)
Private Function LeidoL(ByVal nombre As String, ByVal defecto As Long) As Long
    Dim v As Double
    v = Leido(nombre, defecto)
    If v < -1000000# Or v > 1000000# Then LeidoL = defecto Else LeidoL = CLng(v)
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
