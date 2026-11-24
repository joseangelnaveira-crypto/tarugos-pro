Attribute VB_Name = "mPosiciones"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mPosiciones - FASE 2: numero y posicion de los tarugos de cada union.
'
'   N = N inicial + IncN * ENTERO((L - Lref) / IncLong)     si L > Lref
'   N = N inicial                                           si L <= Lref
' Extremos a DistExtremos del borde de la union; reparto uniforme y simetrico.
'
' Uniones MAS CORTAS que la longitud de referencia: N proporcional a la longitud
'   N = redondeo( N inicial * L / Lref ), minimo 1   (y 1 si L < incremento de longitud)
' Ejemplo con Lref = 1000 y N inicial = 3:
'   L (mm)   300  500  700  1000  1300  1600  1900  2200
'   N          1    2    2     3     4     5     6     7        (valores por defecto)
' El usuario puede FIJAR el numero de una union concreta (NFijos), que prevalece sobre la regla.
'------------------------------------------------------------------------------
'==============================================================================

' Aviso del ultimo CalculaPosiciones cuando se han puesto MENOS tarugos de los del reparto
' estandar por ser la union corta ("" si no ha hecho falta). mUniones lo anade al motivo.
Public NotaReparto As String

' Numero de tarugos fijado a mano por union (clave = ClaveUnion). Lo pone la ventana de revision;
' se conserva al RECALCULAR y se borra al empezar un analisis nuevo (AUTOMATICO / MANUAL).
Public NFijos As Collection

' Identifica una union de forma estable entre analisis: piezas y origen de la linea de tarugos (mm)
Public Function ClaveUnion(un As TUnion) As String
    ClaveUnion = Piezas(un.P1).Nombre & "|" & Piezas(un.P2).Nombre & "|" & _
                 Format(un.Origen.X * 1000, "0.0") & ";" & Format(un.Origen.Y * 1000, "0.0") & ";" & Format(un.Origen.Z * 1000, "0.0")
End Function

Public Sub FijaN(ByVal clave As String, ByVal n As Long)
    If NFijos Is Nothing Then Set NFijos = New Collection
    On Error Resume Next
    NFijos.Remove clave
    On Error GoTo 0
    If n > 0 Then NFijos.Add n, clave
End Sub

' Numero fijado por el usuario para esa union (0 = ninguno)
Public Function NFijado(ByVal clave As String) As Long
    If NFijos Is Nothing Then Exit Function
    On Error Resume Next
    NFijado = CLng(NFijos(clave))
End Function

Public Sub BorraNFijos()
    Set NFijos = Nothing
End Sub

Public Function NumTarugos(ByVal Lmm As Double) As Long
    ' Regla del usuario: una union (tramo entre huecos) mas corta que el incremento de longitud
    ' (300 mm por defecto) lleva UN solo tarugo, centrado.
    If Lmm < Cfg.IncLong Then NumTarugos = 1: Exit Function
    If Lmm <= Cfg.LongRef Then
        ' proporcional a la longitud (0,0001 de tolerancia: 500 mm medidos como 499,99 siguen dando 2)
        NumTarugos = Int(Cfg.NInicial * Lmm / Cfg.LongRef + 0.5001)
        If NumTarugos < 1 Then NumTarugos = 1
        If NumTarugos > Cfg.NInicial Then NumTarugos = Cfg.NInicial
    Else
        NumTarugos = Cfg.NInicial + Cfg.IncN * Int((Lmm - Cfg.LongRef) / Cfg.IncLong)
    End If
End Function

' "" si caben n tarugos en la union de Lmm mm; si no, el motivo. sep = separacion resultante (n > 1).
' n = 1 va centrado y solo exige una longitud minima (no usa la distancia a los extremos).
Private Function CabenN(ByVal Lmm As Double, ByVal d As Double, ByVal n As Long, ByRef sep As Double) As String
    Dim util As Double
    sep = 0
    If n = 1 Then
        If Lmm < 2 * Cfg.Diametro Then CabenN = "Union demasiado corta (" & Format(Lmm, "0.0") & " mm)"
    Else
        util = Lmm - 2 * d
        If util <= 0 Then
            CabenN = "Longitud util insuficiente: " & Format(Lmm, "0.0") & " mm con " & d & " mm a cada extremo"
            Exit Function
        End If
        sep = util / (n - 1)
        ' (tolerancia de 0,001 mm: una union de 200,0000 mm medida como 199,9999 no debe perder un tarugo)
        If sep < Cfg.SepMinFactor * Cfg.Diametro - 0.001 Then
            CabenN = "Separacion entre tarugos (" & Format(sep, "0.0") & " mm) menor que la minima (" & _
                     Format(Cfg.SepMinFactor * Cfg.Diametro, "0.0") & " mm)"
        End If
    End If
End Function

' Calcula un.NT y un.Pos(). Devuelve "" si todo va bien o el motivo del fallo
' (en ese caso NT = 0 y no se crea ningun tarugo).
Public Function CalculaPosiciones(un As TUnion) As String
    un.NT = 0
    Erase un.Pos
    Erase un.Activo
    Erase un.SinMat

    NotaReparto = ""
    Dim Lmm As Double, d As Double, n As Long, n0 As Long
    Lmm = un.L / MM
    d = Cfg.DistExtremos
    n0 = NumTarugos(Lmm)
    Dim nFix As Long
    nFix = NFijado(ClaveUnion(un))
    If nFix > 0 Then n0 = nFix
    If n0 < 1 Then CalculaPosiciones = "Numero de tarugos invalido": Exit Function

    ' Si no caben los n0 tarugos del reparto estandar (union corta), se prueba con uno menos,
    ' y asi hasta 1 tarugo centrado. Cuando n0 SI cabe, el resultado es exactamente el de siempre.
    Dim sep As Double, motivo As String
    n = n0
    Do
        motivo = CabenN(Lmm, d, n, sep)
        If motivo = "" Then Exit Do
        If n <= 1 Then CalculaPosiciones = motivo: Exit Function
        n = n - 1
    Loop
    If nFix > 0 Then
        NotaReparto = "Numero fijado a mano: " & nFix & IIf(n < nFix, " (solo caben " & n & ")", "")
    ElseIf n0 = 1 And Cfg.NInicial > 1 Then
        NotaReparto = "Union de " & Format(Lmm, "0") & " mm (menos de " & Format(Cfg.IncLong, "0") & "): 1 tarugo centrado"
    ElseIf n < n0 Then
        NotaReparto = "Union corta (" & Format(Lmm, "0") & " mm): " & n & " tarugo" & IIf(n = 1, "", "s") & _
                      " en lugar de " & n0 & IIf(n = 1, " (centrado)", "")
    End If

    ReDim un.Pos(0 To n - 1)
    ReDim un.Activo(0 To n - 1)
    ReDim un.SinMat(0 To n - 1)
    Dim k As Long, s As Double
    For k = 0 To n - 1
        If n = 1 Then s = Lmm / 2 Else s = d + k * sep
        s = s * MM
        un.Pos(k) = V3(un.Origen.X + un.Dir.X * s, un.Origen.Y + un.Dir.Y * s, un.Origen.Z + un.Dir.Z * s)
        un.Activo(k) = True
    Next k
    un.NT = n
End Function
