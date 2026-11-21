Attribute VB_Name = "mPosiciones"
Option Explicit
'==============================================================================
' mPosiciones - FASE 2: numero y posicion de los tarugos de cada union.
'
'   N = N inicial + IncN * ENTERO((L - Lref) / IncLong)     si L > Lref
'   N = N inicial                                           si L <= Lref
' Extremos a DistExtremos del borde de la union; reparto uniforme y simetrico.
'------------------------------------------------------------------------------
'   L (mm)    600 1000 1200 1300 1600 1900 2200
'   N          3    3    3    4    5    6    7     (valores por defecto)
'==============================================================================

Public Function NumTarugos(ByVal Lmm As Double) As Long
    If Lmm <= Cfg.LongRef Then
        NumTarugos = Cfg.NInicial
    Else
        NumTarugos = Cfg.NInicial + Cfg.IncN * Int((Lmm - Cfg.LongRef) / Cfg.IncLong)
    End If
End Function

' Calcula un.NT y un.Pos(). Devuelve "" si todo va bien o el motivo del fallo
' (en ese caso NT = 0 y no se crea ningun tarugo).
Public Function CalculaPosiciones(un As TUnion) As String
    un.NT = 0
    Erase un.Pos
    Erase un.Activo

    Dim Lmm As Double, d As Double, n As Long
    Lmm = un.L / MM
    d = Cfg.DistExtremos
    n = NumTarugos(Lmm)
    If n < 1 Then CalculaPosiciones = "Numero de tarugos invalido": Exit Function

    Dim util As Double, sep As Double
    If n = 1 Then
        If Lmm < 2 * Cfg.Diametro Then CalculaPosiciones = "Union demasiado corta (" & Format(Lmm, "0.0") & " mm)": Exit Function
        sep = 0
    Else
        util = Lmm - 2 * d
        If util <= 0 Then
            CalculaPosiciones = "Longitud util insuficiente: " & Format(Lmm, "0.0") & " mm con " & d & " mm a cada extremo"
            Exit Function
        End If
        sep = util / (n - 1)
        If sep < Cfg.SepMinFactor * Cfg.Diametro Then
            CalculaPosiciones = "Separacion entre tarugos (" & Format(sep, "0.0") & " mm) menor que la minima (" & _
                                Format(Cfg.SepMinFactor * Cfg.Diametro, "0.0") & " mm)"
            Exit Function
        End If
    End If

    ReDim un.Pos(0 To n - 1)
    ReDim un.Activo(0 To n - 1)
    Dim k As Long, s As Double
    For k = 0 To n - 1
        If n = 1 Then s = Lmm / 2 Else s = d + k * sep
        s = s * MM
        un.Pos(k) = V3(un.Origen.X + un.Dir.X * s, un.Origen.Y + un.Dir.Y * s, un.Origen.Z + un.Dir.Z * s)
        un.Activo(k) = True
    Next k
    un.NT = n
End Function
