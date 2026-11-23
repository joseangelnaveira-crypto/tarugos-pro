Attribute VB_Name = "mUniones"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mUniones - FASE 2: deteccion de uniones entre piezas.
'
' Una UNION es una zona de contacto entre dos caras planas paralelas y opuestas
' (normales contrarias, mismo plano) cuyas regiones se solapan.
'   - Cara "canto": cara de la pieza que NO es una de sus dos caras grandes.
'   - Union canto-cara  : un canto apoya en la cara grande de otra pieza.
'                         Pieza 1 = la del canto (profundidad grande), Pieza 2 = la de la cara.
'   - Union canto-canto : ambos cantos -> roles ambiguos -> DUDOSA (revisar).
'   - Union cara-cara   : tableros apilados -> DESCARTADA (no es union de tarugos).
' Estados: VALIDA / DUDOSA / DESCARTADA. Solo las VALIDAS generan tarugos.
' Longitudes internas en metros.
'==============================================================================

Public Type TUnion
    Id As String
    P1 As Long               ' indice de pieza 1 (recibe Prof1)
    P2 As Long               ' indice de pieza 2 (recibe Prof2)
    Tipo As String           ' CANTO-CARA / CANTO-CANTO / CARA-CARA / OTRA
    Estado As String         ' VALIDA / DUDOSA / DESCARTADA
    Motivo As String
    N1 As TV3                ' normal saliente del contacto en la pieza 1
    Dir As TV3               ' direccion de la union (eje de reparto)
    Origen As TV3            ' inicio de la linea de tarugos (plano de contacto, centrada)
    L As Double              ' longitud de la union (m)
    W As Double              ' anchura del contacto (m)
    NT As Long               ' numero de tarugos
    Pos() As TV3             ' centro de cada tarugo en el plano de contacto
    Activo() As Boolean
    SinMat() As Boolean      ' tarugo descartado por caer en un hueco de una de las piezas (mHuecos)
    Hecho() As Long          ' por tarugo: 0 = nada hecho, 1 = agujero ya en la pieza 1, 2 = ya en la pieza 2, 3 = en las dos
    Cerca() As Boolean       ' tarugo descartado por haber un agujero ya hecho a menos de la separacion minima (tambien marca SinMat)
End Type

Public Uniones() As TUnion
Public NUniones As Long
Private motivoVert As String            ' por que VerticesCara devolvio False

'------------------------------------------------------------------------------
' Recorre las parejas de piezas candidatas y registra las uniones
'------------------------------------------------------------------------------
Public Sub DetectaUniones()
    NUniones = 0
    Erase Uniones
    Dim a As Long, b As Long
    mOptim.Inicia Application.SldWorks, "Tarugos: buscando uniones", NPiezas
    For a = 0 To NPiezas - 2
        If mOptim.Paso(a) Then Exit For
        If Piezas(a).Estado = "OK" Then
            For b = a + 1 To NPiezas - 1
                If Piezas(b).Estado = "OK" Then
                    ' filtro barato por cajas envolventes: solo las parejas que se tocan
                    ' (o casi) llegan a comparar caras
                    If CajasSeTocan(Piezas(a), Piezas(b), 0.001) Then
                        ComparaPareja a, b
                    End If
                End If
            Next b
        End If
    Next a
    mOptim.Fin

    Dim i As Long
    For i = 0 To NUniones - 1
        Uniones(i).Id = "UNION_" & Format(i + 1, "000")
    Next i
End Sub

Private Function CajasSeTocan(pa As TPieza, pb As TPieza, ByVal margen As Double) As Boolean
    CajasSeTocan = Not (pa.BMax.X + margen < pb.BMin.X Or pb.BMax.X + margen < pa.BMin.X Or _
                        pa.BMax.Y + margen < pb.BMin.Y Or pb.BMax.Y + margen < pa.BMin.Y Or _
                        pa.BMax.Z + margen < pb.BMin.Z Or pb.BMax.Z + margen < pa.BMin.Z)
End Function

Private Sub ComparaPareja(a As Long, b As Long)
    Dim i As Long, j As Long, gap As Double
    For i = 0 To Piezas(a).NCaras - 1
        For j = 0 To Piezas(b).NCaras - 1
            If VDot(Piezas(a).Caras(i).N, Piezas(b).Caras(j).N) < -Cfg.TolAngulo Then
                gap = Abs(VDot(VSub(Piezas(b).Caras(j).P, Piezas(a).Caras(i).P), Piezas(a).Caras(i).N))
                If gap <= 0.002 Then          ' hasta 2 mm: contacto o "casi contacto"
                    If CarasCajaSolapa(Piezas(a).Caras(i), Piezas(b).Caras(j), 0.002) Then
                        AnalizaContacto a, i, b, j, gap
                    End If
                End If
            End If
        Next j
    Next i
End Sub

Private Function CarasCajaSolapa(ca As TCara, cb As TCara, ByVal m As Double) As Boolean
    CarasCajaSolapa = Not (ca.BMax.X + m < cb.BMin.X Or cb.BMax.X + m < ca.BMin.X Or _
                           ca.BMax.Y + m < cb.BMin.Y Or cb.BMax.Y + m < ca.BMin.Y Or _
                           ca.BMax.Z + m < cb.BMin.Z Or cb.BMax.Z + m < ca.BMin.Z)
End Function

'------------------------------------------------------------------------------
' Vertices de una cara plana en coordenadas de ensamblaje. False si tiene curvas.
'------------------------------------------------------------------------------
Private Function VerticesCara(f As SldWorks.Face2, pts() As TV3, ByRef n As Long) As Boolean
    n = 0
    motivoVert = ""
    ' Solo el bucle EXTERIOR de la cara: los agujeros existentes (bucles interiores,
    ' aristas circulares) no cambian el contorno de contacto.
    Dim lp As SldWorks.Loop2, lp0 As SldWorks.Loop2
    Set lp0 = f.GetFirstLoop
    Set lp = lp0
    Do While Not lp Is Nothing
        If lp.IsOuter <> 0 Then Exit Do          ' (Boolean/Long segun version: se compara con 0)
        Set lp = lp.GetNext
    Loop
    If lp Is Nothing Then Set lp = lp0            ' si ninguno se marca como exterior, el primero
    If lp Is Nothing Then motivoVert = "cara sin bucles": Exit Function
    Dim edgs As Variant
    edgs = lp.GetEdges
    If IsEmpty(edgs) Then motivoVert = "GetEdges vacio": Exit Function
    ReDim pts(0 To 2 * (UBound(edgs) + 1))
    Dim k As Long, e As SldWorks.Edge, v1 As SldWorks.Vertex, v2 As SldWorks.Vertex
    For k = 0 To UBound(edgs)
        Set e = edgs(k)
        If e.GetCurve Is Nothing Then motivoVert = "arista sin curva": Exit Function
        Set v1 = e.GetStartVertex
        Set v2 = e.GetEndVertex
        If v1 Is Nothing Or v2 Is Nothing Then
            ' una arista CURVA cerrada (circulo) no tiene vertices: se ignora; una recta si debe tenerlos
            If e.GetCurve.Identity = 3001 Then motivoVert = "arista sin vertices": Exit Function
        Else
            ' Las aristas curvas (esquinas redondeadas, "Redondeo") se sustituyen por la cuerda
            ' entre sus extremos: el contorno queda algo mas pequeno que el real (conservador).
            pts(n) = VFromArr(v1.GetPoint): n = n + 1
            pts(n) = VFromArr(v2.GetPoint): n = n + 1
        End If
    Next k
    If n < 3 Then motivoVert = "menos de 3 vertices": Exit Function
    VerticesCara = True
End Function

'------------------------------------------------------------------------------
' Region de contacto entre la cara i de la pieza a y la cara j de la pieza b
'------------------------------------------------------------------------------
Private Sub AnalizaContacto(a As Long, i As Long, b As Long, j As Long, ByVal gap As Double)
    Dim un As TUnion
    un.Tipo = "OTRA": un.Estado = "DESCARTADA"

    Dim ca As TCara, cb As TCara
    ca = Piezas(a).Caras(i)
    cb = Piezas(b).Caras(j)

    ' ---- rol de cada cara ----
    Dim cantoA As Boolean, cantoB As Boolean
    cantoA = (Piezas(a).EsTablero And i <> Piezas(a).iCaraA And i <> Piezas(a).iCaraB)
    cantoB = (Piezas(b).EsTablero And j <> Piezas(b).iCaraA And j <> Piezas(b).iCaraB)

    ' ---- poligonos en el plano de contacto ----
    Dim pa() As TV3, na As Long, pb() As TV3, nb As Long
    Dim okA As Boolean, okB As Boolean, mA As String, mB As String
    okA = VerticesCara(ca.Ref, pa, na): mA = motivoVert
    okB = VerticesCara(cb.Ref, pb, nb): mB = motivoVert
    If Not okA Or Not okB Then
        un.P1 = a: un.P2 = b: un.Tipo = "OTRA": un.Estado = "DUDOSA"
        un.Motivo = "No se pudieron leer los vertices de la cara (" & mA & IIf(mA <> "" And mB <> "", " / ", "") & mB & "): requiere analisis completo"
        AnadeUnion un
        Exit Sub
    End If

    Dim n As TV3, e1 As TV3, e2 As TV3
    n = ca.N
    If Abs(n.X) <= Abs(n.Y) And Abs(n.X) <= Abs(n.Z) Then
        e1 = VUnit(VCross(n, V3(1, 0, 0)))
    ElseIf Abs(n.Y) <= Abs(n.Z) Then
        e1 = VUnit(VCross(n, V3(0, 1, 0)))
    Else
        e1 = VUnit(VCross(n, V3(0, 0, 1)))
    End If
    e2 = VCross(n, e1)

    Dim qa() As TP2, qb() As TP2, k As Long, d As TV3
    ReDim qa(0 To na - 1): ReDim qb(0 To nb - 1)
    For k = 0 To na - 1
        d = VSub(pa(k), ca.P): qa(k).X = VDot(d, e1): qa(k).Y = VDot(d, e2)
    Next k
    For k = 0 To nb - 1
        d = VSub(pb(k), ca.P): qb(k).X = VDot(d, e1): qb(k).Y = VDot(d, e2)
    Next k

    Dim ha() As TP2, nha As Long, hb() As TP2, nhb As Long
    Casco2 qa, na, ha, nha
    Casco2 qb, nb, hb, nhb
    If nha < 3 Or nhb < 3 Then Exit Sub
    ' Cara "compleja": con agujeros (varios bucles) o con contorno exterior no convexo
    ' (muescas, recortes). Solo esas se comprueban punto a punto (mHuecos); las caras
    ' rectangulares se tratan exactamente como siempre.
    Dim complejo As Boolean
    complejo = (ca.Ref.GetLoopCount > 1) Or (cb.Ref.GetLoopCount > 1) Or (nha <> na \ 2) Or (nhb <> nb \ 2)

    Dim ov() As TP2, nov As Long
    Intersecta2 ha, nha, hb, nhb, ov, nov
    If nov < 3 Then Exit Sub
    Dim aov As Double
    aov = Abs(Area2(ov, nov))
    If aov < Cfg.AreaMinCara * MM * MM Then Exit Sub        ' solo se tocan en un borde

    un.P1 = a: un.P2 = b

    ' ---- direccion de la union = arista mas larga del solape ----
    Dim lbest As Double, ibest As Long, l As Double, k2 As Long
    For k = 0 To nov - 1
        k2 = (k + 1) Mod nov
        l = Sqr((ov(k2).X - ov(k).X) ^ 2 + (ov(k2).Y - ov(k).Y) ^ 2)
        If l > lbest Then lbest = l: ibest = k
    Next k
    k2 = (ibest + 1) Mod nov
    Dim ux As Double, uy As Double
    ux = (ov(k2).X - ov(ibest).X) / lbest
    uy = (ov(k2).Y - ov(ibest).Y) / lbest
    Dim u3 As TV3, v3b As TV3
    u3 = V3(e1.X * ux + e2.X * uy, e1.Y * ux + e2.Y * uy, e1.Z * ux + e2.Z * uy)
    v3b = VUnit(VCross(n, u3))

    Dim smin As Double, smax As Double, tmin As Double, tmax As Double, s As Double, t As Double
    smin = 1E+30: tmin = 1E+30: smax = -1E+30: tmax = -1E+30
    For k = 0 To nov - 1
        s = ov(k).X * ux + ov(k).Y * uy
        t = -ov(k).X * uy + ov(k).Y * ux            ' componente perpendicular (n x u en el plano)
        If s < smin Then smin = s
        If s > smax Then smax = s
        If t < tmin Then tmin = t
        If t > tmax Then tmax = t
    Next k
    un.L = smax - smin
    un.W = tmax - tmin
    un.Dir = u3
    ' origen: inicio de la linea central (en el plano de contacto)
    Dim cs As Double, ct As Double
    cs = smin: ct = (tmin + tmax) / 2
    Dim px As Double, py As Double
    px = cs * ux - ct * uy
    py = cs * uy + ct * ux
    un.Origen = V3(ca.P.X + e1.X * px + e2.X * py, ca.P.Y + e1.Y * px + e2.Y * py, ca.P.Z + e1.Z * px + e2.Z * py)

    ' ---- tipo y roles ----
    If Not Piezas(a).EsTablero Or Not Piezas(b).EsTablero Then
        un.Tipo = "OTRA": un.Estado = "DUDOSA"
        un.Motivo = "Una de las piezas no se reconoce como tablero"
        un.N1 = ca.N
    ElseIf cantoA And Not cantoB Then
        un.Tipo = "CANTO-CARA": un.P1 = a: un.P2 = b: un.N1 = ca.N: un.Estado = "VALIDA"
    ElseIf cantoB And Not cantoA Then
        un.Tipo = "CANTO-CARA": un.P1 = b: un.P2 = a: un.N1 = cb.N: un.Estado = "VALIDA"
    ElseIf cantoA And cantoB Then
        un.Tipo = "CANTO-CANTO": un.P1 = a: un.P2 = b: un.N1 = ca.N: un.Estado = "DUDOSA"
        un.Motivo = "Canto con canto: hay que decidir cual es la pieza 1 (profundidad " & Cfg.Prof1 & " mm)"
    Else
        un.Tipo = "CARA-CARA": un.Estado = "DESCARTADA": un.N1 = ca.N
        un.Motivo = "Tableros apilados cara con cara: no es una union de tarugos"
    End If

    ' ---- piezas proximas pero no unidas ----
    If gap > Cfg.Tol * MM Then
        un.Estado = "DESCARTADA"
        un.Motivo = AnadeMotivo(un.Motivo, "Caras paralelas separadas " & Format(gap * 1000, "0.00") & " mm: no estan en contacto")
    End If

    ' ---- comprobaciones de espacio y reparto (solo si se puede perforar) ----
    If un.Estado <> "DESCARTADA" And un.Tipo <> "OTRA" Then
        Dim espesorCanto As Double
        espesorCanto = Piezas(un.P1).Espesor
        If un.Tipo = "CANTO-CANTO" Then espesorCanto = MinD(Piezas(a).Espesor, Piezas(b).Espesor)
        If un.W < 0.8 * espesorCanto Then
            If un.Estado = "VALIDA" Then un.Estado = "DUDOSA"
            un.Motivo = AnadeMotivo(un.Motivo, "Contacto estrecho (" & Format(un.W * 1000, "0.0") & " mm) frente al espesor del canto")
        End If
        If un.W - Cfg.Diametro * MM < 2 * Cfg.MinMaterial * MM Then
            If un.Estado = "VALIDA" Then un.Estado = "DUDOSA"
            un.Motivo = AnadeMotivo(un.Motivo, "Anchura de contacto insuficiente para el diametro")
        End If

        Dim mi As TV3, av1 As Double, av2 As Double
        mi = V3(-un.N1.X, -un.N1.Y, -un.N1.Z)
        av1 = EspacioPieza(Piezas(un.P1), ca.P, mi)
        av2 = EspacioPieza(Piezas(un.P2), ca.P, un.N1)
        If av1 < (Cfg.Prof1 + Cfg.MinMaterial) * MM Then
            If un.Estado = "VALIDA" Then un.Estado = "DUDOSA"
            un.Motivo = AnadeMotivo(un.Motivo, "Pieza 1: material insuficiente para " & Cfg.Prof1 & " mm (hay " & Format(av1 * 1000, "0.0") & ")")
        End If
        If av2 < (Cfg.Prof2 + Cfg.MinMaterial) * MM Then
            If un.Estado = "VALIDA" Then un.Estado = "DUDOSA"
            un.Motivo = AnadeMotivo(un.Motivo, "Pieza 2: material insuficiente para " & Cfg.Prof2 & " mm (hay " & Format(av2 * 1000, "0.0") & ")")
        End If

        ' reparto de tarugos (se calcula tambien para uniones dudosas, para revisarlas)
        Dim errPos As String
        errPos = CalculaPosiciones(un)
        If errPos = "" And mPosiciones.NotaReparto <> "" Then un.Motivo = AnadeMotivo(un.Motivo, mPosiciones.NotaReparto)
        If errPos <> "" Then
            un.Estado = "DUDOSA"
            un.Motivo = AnadeMotivo(un.Motivo, errPos)
        ElseIf complejo And un.NT > 0 Then
            ' comprobar que cada tarugo cae sobre material de las DOS piezas (mover o descartar)
            Dim txtHueco As String
            txtHueco = mHuecos.AjustaPorHuecos(un, ca.Ref, cb.Ref)
            If txtHueco <> "" Then un.Motivo = AnadeMotivo(un.Motivo, txtHueco)
        End If
        If un.NT > 0 Then MarcaHechos un, ca.Ref, cb.Ref, (un.P1 = a)
    End If
    AnadeUnion un
End Sub

Private Function AnadeMotivo(ByVal actual As String, ByVal nuevo As String) As String
    If actual = "" Then AnadeMotivo = nuevo Else AnadeMotivo = actual & "; " & nuevo
End Function

' Distancia disponible en la pieza desde el punto P en el sentido dirIn
' (maximo de la proyeccion de las 8 esquinas de su caja; estimacion para piezas alineadas)
Private Function EspacioPieza(pz As TPieza, P As TV3, dirIn As TV3) As Double
    Dim i As Long, q As TV3, m As Double, d As Double
    m = -1E+30
    For i = 0 To 7
        q = V3(IIf(i And 1, pz.BMax.X, pz.BMin.X), IIf(i And 2, pz.BMax.Y, pz.BMin.Y), IIf(i And 4, pz.BMax.Z, pz.BMin.Z))
        d = VDot(VSub(q, P), dirIn)
        If d > m Then m = d
    Next i
    EspacioPieza = m
End Function

Private Sub AnadeUnion(un As TUnion)
    ReDim Preserve Uniones(0 To NUniones)
    Uniones(NUniones) = un
    NUniones = NUniones + 1
End Sub

' Marca los tarugos cuyo agujero YA existe en el modelo (el mueble ya trae agujeros de tarugo).
' fa/fb son las caras de contacto de las piezas a/b; aEsP1 indica si a es la pieza 1.
Private Sub MarcaHechos(un As TUnion, fa As SldWorks.Face2, fb As SldWorks.Face2, ByVal aEsP1 As Boolean)
    Dim k As Long, ha As Boolean, hb As Boolean
    ReDim un.Hecho(0 To un.NT - 1)
    For k = 0 To un.NT - 1
        ha = AgujeroEnCara(fa, un.Pos(k))
        hb = AgujeroEnCara(fb, un.Pos(k))
        If aEsP1 Then
            un.Hecho(k) = IIf(ha, 1, 0) + IIf(hb, 2, 0)
        Else
            un.Hecho(k) = IIf(hb, 1, 0) + IIf(ha, 2, 0)
        End If
    Next k

    ' Un agujero de tarugo ya hecho a menos de la separacion minima (50 mm por defecto) de un
    ' tarugo NUEVO: ese tarugo se descarta (en las dos piezas), como si cayera en un hueco.
    ' Los que coinciden con un agujero hecho (Hecho > 0) se quedan: se saltan al mecanizar.
    ReDim un.Cerca(0 To un.NT - 1)
    Dim dMin As Double, txt As String
    dMin = Cfg.SepMinFactor * Cfg.Diametro * MM - 0.000001
    For k = 0 To un.NT - 1
        If un.Hecho(k) = 0 And Not un.SinMat(k) Then
            If AgujeroCerca(fa, un.Pos(k), dMin) Or AgujeroCerca(fb, un.Pos(k), dMin) Then
                un.Cerca(k) = True
                un.SinMat(k) = True
                un.Activo(k) = False
                If txt <> "" Then txt = txt & ", "
                txt = txt & (k + 1)
            End If
        End If
    Next k
    If txt <> "" Then
        un.Motivo = AnadeMotivo(un.Motivo, "tarugo(s) " & txt & " descartado(s): hay un agujero ya hecho a menos de " & _
                                Format(Cfg.SepMinFactor * Cfg.Diametro, "0") & " mm")
    End If
End Sub

' True si la cara ya tiene un agujero circular del diametro del tarugo centrado en pp
' (p. ej. los agujeros de tarugo que el mueble ya trae).
Public Function AgujeroEnCara(f As SldWorks.Face2, pp As TV3) As Boolean
    AgujeroEnCara = AgujeroCerca(f, pp, 0.0003)
End Function

' True si la cara tiene un agujero circular del diametro del tarugo con el centro a menos
' de dMax (m) de pp. Solo se miran los bucles interiores.
Private Function AgujeroCerca(f As SldWorks.Face2, pp As TV3, ByVal dMax As Double) As Boolean
    Dim lp As SldWorks.Loop2, edgs As Variant, k As Long, e As SldWorks.Edge
    Dim cv As SldWorks.Curve, cp As Variant
    On Error Resume Next
    Set lp = f.GetFirstLoop
    Do While Not lp Is Nothing
        If lp.IsOuter = 0 Then
            edgs = lp.GetEdges
            If Not IsEmpty(edgs) Then
                For k = 0 To UBound(edgs)
                    Set e = edgs(k)
                    Set cv = e.GetCurve
                    If Not cv Is Nothing Then
                        If cv.Identity = 3002 Then
                            cp = cv.CircleParams                  ' centro(0..2), eje(3..5), radio(6)
                            If IsArray(cp) Then
                                If Abs(cp(6) - Cfg.Diametro * MM / 2) < 0.00005 Then
                                    If VLen(V3(cp(0) - pp.X, cp(1) - pp.Y, cp(2) - pp.Z)) < dMax Then
                                        AgujeroCerca = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next k
            End If
        End If
        Set lp = lp.GetNext
    Loop
End Function
