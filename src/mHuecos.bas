Attribute VB_Name = "mHuecos"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mHuecos - comprueba que cada tarugo cae sobre MATERIAL real en las dos piezas.
'
' El reparto de mPosiciones se calcula con el solape de los contornos exteriores
' RELLENOS (casco convexo): no sabe de agujeros, muescas ni recortes de las caras de
' contacto. Sin esta comprobacion un tarugo podia caer en un hueco de una pieza: se
' taladraba la otra y la del hueco quedaba sin su agujero (o sin guardar).
'
' Solo se llama para uniones "complejas" (alguna cara con agujeros o con contorno no
' convexo). Las caras rectangulares NO pasan por aqui: se tratan exactamente como antes.
'
' Regla acordada con el usuario:
'   1. Si el tarugo cae en un hueco, o a menos de DistExtremos del borde del hueco,
'      se busca el punto mas cercano de la linea que tenga DistExtremos de material a
'      cada lado y respete la separacion minima con los demas tarugos.
'   2. Si no hay ninguno, el tarugo se DESCARTA (en las dos piezas) y queda marcado
'      como "sin material" (SinMat) en la revision.
'
' "Hay material" = el punto esta sobre la cara (distancia a su punto mas cercano,
' Face2.GetClosestPointOn, menor que TOL_CARA) en AMBAS caras de contacto.
' Los huecos muy pequenos (un agujero de tarugo, de diametro parecido al de la macro)
' se ignoran: son los taladros que esta misma macro ya hizo en una ejecucion anterior
' y no deben desplazar los tarugos al volver a ejecutarla.
'==============================================================================

Private Const PASO As Double = 0.005          ' 5 mm entre muestras a lo largo de la linea de tarugos
Private Const TOL_CARA As Double = 0.0005     ' 0,5 mm

' True si el punto p esta sobre la cara f. Si la API no responde se asume que hay
' material (no se cambia nada: el fallo es "no hacer nada", nunca mover tarugos a ciegas).
Private Function EnCara(f As SldWorks.Face2, p As TV3) As Boolean
    Dim c As Variant, q As TV3
    EnCara = True
    On Error Resume Next
    c = f.GetClosestPointOn(p.X, p.Y, p.Z)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If Not IsArray(c) Then Exit Function
    q = V3(c(0), c(1), c(2))
    EnCara = (VLen(VSub(q, p)) <= TOL_CARA)
End Function

' Hay material en el punto de la linea a distancia s del origen, y a +-radio del taladro
' en el sentido transversal (tv), en las dos caras
Private Function HayMaterial(un As TUnion, ByVal s As Double, tv As TV3, _
                             fa As SldWorks.Face2, fb As SldWorks.Face2) As Boolean
    Dim p As TV3, q As TV3, r As Double
    r = Cfg.Diametro * MM / 2
    p = V3(un.Origen.X + un.Dir.X * s, un.Origen.Y + un.Dir.Y * s, un.Origen.Z + un.Dir.Z * s)
    If Not EnCara(fa, p) Then Exit Function
    If Not EnCara(fb, p) Then Exit Function
    q = V3(p.X + tv.X * r, p.Y + tv.Y * r, p.Z + tv.Z * r)
    If Not EnCara(fa, q) Then Exit Function
    If Not EnCara(fb, q) Then Exit Function
    q = V3(p.X - tv.X * r, p.Y - tv.Y * r, p.Z - tv.Z * r)
    If Not EnCara(fa, q) Then Exit Function
    If Not EnCara(fb, q) Then Exit Function
    HayMaterial = True
End Function

' Ajusta un.Pos / un.Activo / un.SinMat. Devuelve "" si no ha cambiado nada, o el texto
' (para el Motivo de la union) de lo que ha movido o descartado.
Public Function AjustaPorHuecos(un As TUnion, fa As SldWorks.Face2, fb As SldWorks.Face2) As String
    If un.NT < 1 Or un.L <= 0 Then Exit Function

    Dim tv As TV3
    tv = VUnit(VCross(un.N1, un.Dir))

    Dim ns As Long, i As Long, s As Double
    ns = Int(un.L / PASO) + 1                 ' muestras 0..ns (la ultima se recorta a s = L)
    Dim ok() As Boolean
    ReDim ok(0 To ns)
    For i = 0 To ns
        s = i * PASO
        If s > un.L Then s = un.L
        ok(i) = HayMaterial(un, s, tv, fa, fb)
    Next i

    ' ignorar huecos pequenos (taladros de tarugo ya hechos): rachas de pocas muestras
    Dim racha As Long, j As Long, hayHueco As Boolean
    i = 0
    Do While i <= ns
        If ok(i) Then
            i = i + 1
        Else
            j = i
            Do While j <= ns
                If ok(j) Then Exit Do
                j = j + 1
            Loop
            racha = j - i                      ' muestras seguidas sin material
            If (racha + 1) * PASO <= 2 * Cfg.Diametro * MM Then
                Dim q As Long
                For q = i To j - 1: ok(q) = True: Next q
            Else
                hayHueco = True
            End If
            i = j
        End If
    Loop
    If Not hayHueco Then Exit Function         ' todo con material: no se toca nada

    ' mal(i) = muestras sin material acumuladas hasta i (para comprobar ventanas en O(1))
    Dim mal() As Long
    ReDim mal(0 To ns)
    For i = 0 To ns
        If i > 0 Then mal(i) = mal(i - 1)
        If Not ok(i) Then mal(i) = mal(i) + 1
    Next i

    Dim d As Double, sepMin As Double
    d = Cfg.DistExtremos * MM
    sepMin = Cfg.SepMinFactor * Cfg.Diametro * MM

    Dim n As Long, k As Long
    n = un.NT
    Dim sFin() As Double, sIdeal() As Double
    ReDim sFin(0 To n - 1): ReDim sIdeal(0 To n - 1)
    For k = 0 To n - 1
        sIdeal(k) = VDot(VSub(un.Pos(k), un.Origen), un.Dir)
        sFin(k) = sIdeal(k)
    Next k

    Dim txt As String, mejor As Double, dist As Double, sC As Double, hallado As Boolean
    Dim iMin As Long, iMax As Long, cand As Long
    For k = 0 To n - 1
        If Not LibreDeHueco(sIdeal(k), d, ns, ok, mal) Then
            hallado = False: mejor = 1E+30
            iMin = -Int(-d / PASO)                       ' techo
            iMax = Int((un.L - d) / PASO)                ' suelo
            For cand = iMin To iMax
                sC = cand * PASO
                If cand >= 0 And cand <= ns Then
                    dist = Abs(sC - sIdeal(k))
                    If dist < mejor Then
                        If LibreDeHueco(sC, d, ns, ok, mal) Then
                            If SeparadoDeOtros(sC, k, n, sFin, sIdeal, un, sepMin) Then
                                mejor = dist: hallado = True: s = sC
                            End If
                        End If
                    End If
                End If
            Next cand
            If hallado Then
                sFin(k) = s
                un.Pos(k) = V3(un.Origen.X + un.Dir.X * s, un.Origen.Y + un.Dir.Y * s, un.Origen.Z + un.Dir.Z * s)
                txt = Unir(txt, "tarugo " & (k + 1) & " movido de " & Format(sIdeal(k) * 1000, "0") & " a " & _
                           Format(s * 1000, "0") & " mm por un hueco")
            Else
                un.SinMat(k) = True
                un.Activo(k) = False
                txt = Unir(txt, "tarugo " & (k + 1) & " sin material (hueco): descartado en las dos piezas")
            End If
        End If
    Next k
    AjustaPorHuecos = txt
End Function

' La ventana [s-d, s+d] no tiene ninguna muestra sin material, y el propio punto tiene material
Private Function LibreDeHueco(ByVal s As Double, ByVal d As Double, ByVal ns As Long, _
                              ok() As Boolean, mal() As Long) As Boolean
    Dim a As Long, b As Long, previos As Long
    a = Int((s - d) / PASO)
    b = -Int(-(s + d) / PASO)
    If a < 0 Then a = 0
    If b > ns Then b = ns
    If a > b Then Exit Function
    Dim ic As Long
    ic = CLng(s / PASO)
    If ic < 0 Then ic = 0
    If ic > ns Then ic = ns
    If Not ok(ic) Then Exit Function
    If a > 0 Then previos = mal(a - 1)
    LibreDeHueco = ((mal(b) - previos) = 0)
End Function

' Separacion minima con los demas tarugos (los ya colocados en su sitio final, los
' pendientes en su sitio ideal). Los descartados no cuentan.
Private Function SeparadoDeOtros(ByVal sC As Double, ByVal k As Long, ByVal n As Long, _
                                 sFin() As Double, sIdeal() As Double, un As TUnion, _
                                 ByVal sepMin As Double) As Boolean
    Dim j As Long, sj As Double
    For j = 0 To n - 1
        If j <> k Then
            If Not un.SinMat(j) Then
                If j < k Then sj = sFin(j) Else sj = sIdeal(j)
                If Abs(sC - sj) < sepMin Then Exit Function
            End If
        End If
    Next j
    SeparadoDeOtros = True
End Function

Private Function Unir(ByVal a As String, ByVal b As String) As String
    If a = "" Then Unir = b Else Unir = a & "; " & b
End Function
