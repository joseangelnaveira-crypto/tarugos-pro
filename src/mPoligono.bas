Attribute VB_Name = "mPoligono"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mPoligono - Geometria 2D minima: casco convexo, area, interseccion de
' poligonos convexos (Sutherland-Hodgman). Sin llamadas a SolidWorks.
' Los poligonos se tratan como CONVEXOS y en sentido antihorario (CCW).
'==============================================================================

Public Type TP2
    X As Double
    Y As Double
End Type

Private Const EPS As Double = 0.0000001      ' 0,1 micras (en metros)

' Area con signo (positiva = antihorario)
Public Function Area2(p() As TP2, ByVal n As Long) As Double
    Dim i As Long, j As Long, s As Double
    For i = 0 To n - 1
        j = (i + 1) Mod n
        s = s + p(i).X * p(j).Y - p(j).X * p(i).Y
    Next i
    Area2 = s / 2
End Function

Private Function Cruz(a As TP2, b As TP2, c As TP2) As Double
    Cruz = (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X)
End Function

' Casco convexo (cadena monotona). Devuelve puntos en sentido antihorario.
Public Sub Casco2(pts() As TP2, ByVal n As Long, hull() As TP2, ByRef nh As Long)
    nh = 0
    If n < 3 Then Exit Sub
    Dim a() As TP2, i As Long, j As Long, t As TP2
    ReDim a(0 To n - 1)
    For i = 0 To n - 1: a(i) = pts(i): Next i
    ' ordenar por X, luego Y (insercion; n pequeno)
    For i = 1 To n - 1
        t = a(i)
        j = i - 1
        Do While j >= 0
            If a(j).X > t.X Or (a(j).X = t.X And a(j).Y > t.Y) Then
                a(j + 1) = a(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        a(j + 1) = t
    Next i
    ' quitar duplicados
    Dim u() As TP2, nu As Long
    ReDim u(0 To n - 1)
    For i = 0 To n - 1
        If nu = 0 Then
            u(0) = a(i): nu = 1
        ElseIf Abs(a(i).X - u(nu - 1).X) > EPS Or Abs(a(i).Y - u(nu - 1).Y) > EPS Then
            u(nu) = a(i): nu = nu + 1
        End If
    Next i
    If nu < 3 Then Exit Sub

    Dim h() As TP2, k As Long
    ReDim h(0 To 2 * nu)
    For i = 0 To nu - 1
        Do While k >= 2
            If Cruz(h(k - 2), h(k - 1), u(i)) <= 0 Then k = k - 1 Else Exit Do
        Loop
        h(k) = u(i): k = k + 1
    Next i
    Dim lim As Long
    lim = k + 1
    For i = nu - 2 To 0 Step -1
        Do While k >= lim
            If Cruz(h(k - 2), h(k - 1), u(i)) <= 0 Then k = k - 1 Else Exit Do
        Loop
        h(k) = u(i): k = k + 1
    Next i
    nh = k - 1
    ReDim hull(0 To nh - 1)
    For i = 0 To nh - 1: hull(i) = h(i): Next i
End Sub

' Interseccion de dos poligonos convexos CCW. out(0..no-1); no=0 si no se tocan.
Public Sub Intersecta2(s() As TP2, ByVal ns As Long, c() As TP2, ByVal nc As Long, _
                       out() As TP2, ByRef no As Long)
    Dim cur() As TP2, nxt() As TP2, ncur As Long, nnxt As Long
    Dim i As Long, k As Long
    ReDim cur(0 To ns + nc + 4)
    For i = 0 To ns - 1: cur(i) = s(i): Next i
    ncur = ns

    Dim a As TP2, b As TP2, p As TP2, q As TP2, dp As Double, dq As Double, t As Double, r As TP2
    For k = 0 To nc - 1
        a = c(k): b = c((k + 1) Mod nc)
        ReDim nxt(0 To ncur + nc + 4)
        nnxt = 0
        For i = 0 To ncur - 1
            q = cur(i)
            p = cur((i + ncur - 1) Mod ncur)
            dq = Cruz(a, b, q)
            dp = Cruz(a, b, p)
            If dq >= -0.000000000001 Then
                If dp < -0.000000000001 Then
                    t = dp / (dp - dq)
                    r.X = p.X + t * (q.X - p.X): r.Y = p.Y + t * (q.Y - p.Y)
                    nxt(nnxt) = r: nnxt = nnxt + 1
                End If
                nxt(nnxt) = q: nnxt = nnxt + 1
            ElseIf dp >= -0.000000000001 Then
                t = dp / (dp - dq)
                r.X = p.X + t * (q.X - p.X): r.Y = p.Y + t * (q.Y - p.Y)
                nxt(nnxt) = r: nnxt = nnxt + 1
            End If
        Next i
        ncur = nnxt
        ReDim cur(0 To ncur + nc + 4)
        For i = 0 To ncur - 1: cur(i) = nxt(i): Next i
        If ncur = 0 Then Exit For
    Next k

    no = ncur
    If no > 0 Then
        ReDim out(0 To no - 1)
        For i = 0 To no - 1: out(i) = cur(i): Next i
    End If
End Sub
