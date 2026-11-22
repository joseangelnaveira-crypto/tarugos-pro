Attribute VB_Name = "mGeometria"
Option Explicit
'==============================================================================
' mGeometria - FASE 1: reconocimiento de piezas y geometria del ensamblaje.
'
' Para cada componente de pieza (no suprimido, resuelto, visible) extrae:
'   - caja envolvente en coordenadas de ensamblaje
'   - caras planas (normal saliente, punto, area, caja) en coordenadas de ensamblaje
'   - si es un "tablero": dos caras grandes paralelas y opuestas -> espesor, largo, ancho
' Todas las longitudes internas en METROS (unidades de la API).
'
' Solo se usan funciones presentes en la API de SOLIDWORKS 2024 (verificadas
' contra sldworks.tlb) para mantener compatibilidad con 2024/2025/2026.
'==============================================================================

Public Type TCara
    N As TV3                 ' normal saliente del material (ensamblaje)
    P As TV3                 ' un punto del plano (ensamblaje)
    Area As Double           ' m2
    BMin As TV3              ' caja alineada a ejes (ensamblaje)
    BMax As TV3
    Ref As SldWorks.Face2
End Type

Public Type TPieza
    Comp As SldWorks.Component2
    Nombre As String
    Ruta As String
    Config As String
    Estado As String         ' OK / SUPRIMIDA / LIGERA / OCULTA / ENSAMBLAJE / SIN GEOMETRIA / ERROR
    BMin As TV3              ' caja del componente (ensamblaje)
    BMax As TV3
    NCaras As Long
    Caras() As TCara
    EsTablero As Boolean
    iCaraA As Long           ' indice cara grande 1
    iCaraB As Long           ' indice cara grande opuesta
    Espesor As Double        ' m
    Largo As Double          ' m
    Ancho As Double          ' m
    ContornoCurvo As Boolean
    Aviso As String
    EnEns As Boolean
    TrDatos As String
    Cuerpo As SldWorks.Body2     ' solo en piezas MULTICUERPO: el cuerpo solido que ES la pieza (Comp = Nothing)
    NombreCuerpo As String       ' nombre del cuerpo al analizar ("" en ensamblajes). OJO: SolidWorks lo
                                 ' renombra al cortarlo; para localizarlo despues se usa su CAJA, no el nombre
End Type

Public Piezas() As TPieza
Public NPiezas As Long
Public Ligeras As Long       ' componentes ligeros que no se han podido analizar

Private mu As SldWorks.MathUtility
' Mantiene vivos los cuerpos (y por tanto las caras guardadas en TCara.Ref).
' Sin esto, SolidWorks libera las caras al salir de AnalizaComponente.
Private Keep As Collection

Private swApp As SldWorks.SldWorks

'------------------------------------------------------------------------------
' Punto de entrada: analiza todos los componentes del ensamblaje activo
'------------------------------------------------------------------------------
Public Sub AnalizaPiezas(app As SldWorks.SldWorks, asm As SldWorks.AssemblyDoc, _
                         Optional soloSeleccion As Boolean = False)
    Set swApp = app
    Set mu = swApp.GetMathUtility
    Set Keep = New Collection
    NPiezas = 0
    Ligeras = 0
    Erase Piezas

    Dim comps As Variant
    comps = asm.GetComponents(False)          ' todos los niveles
    If IsEmpty(comps) Then Exit Sub

    Dim i As Long, c As SldWorks.Component2
    ReDim Piezas(0 To UBound(comps))          ' tamano exacto de una vez (sin ReDim Preserve por pieza)
    mOptim.Inicia app, "Tarugos: analizando piezas", UBound(comps) + 1
    For i = 0 To UBound(comps)
        If mOptim.Paso(i) Then Exit For       ' ESC = cancelar (no se ha modificado nada)
        Set c = comps(i)
        AnalizaComponente c, Piezas(NPiezas)
        NPiezas = NPiezas + 1
    Next i
    mOptim.TiempoPiezas = mOptim.Fin
End Sub

'------------------------------------------------------------------------------
' Pieza MULTICUERPO abierta (no es un ensamblaje): cada cuerpo solido visible es una
' "pieza". Las coordenadas del documento ya son las del "ensamblaje" (transformacion
' identidad, caso que AnalizaComponente ya contempla). Comp = Nothing.
'------------------------------------------------------------------------------
Public Sub AnalizaPiezasDoc(app As SldWorks.SldWorks, md As SldWorks.ModelDoc2)
    Set swApp = app
    Set mu = swApp.GetMathUtility
    Set Keep = New Collection
    NPiezas = 0
    Ligeras = 0
    Erase Piezas

    Dim pd As SldWorks.PartDoc
    Set pd = md
    Dim bodies As Variant
    bodies = pd.GetBodies2(swSolidBody, True)
    If IsEmpty(bodies) Then Exit Sub
    Dim i As Long
    ReDim Piezas(0 To UBound(bodies))
    mOptim.Inicia app, "Tarugos: analizando cuerpos", UBound(bodies) + 1
    For i = 0 To UBound(bodies)
        If mOptim.Paso(i) Then Exit For
        AnalizaCuerpoDoc bodies(i), md, Piezas(NPiezas)
        NPiezas = NPiezas + 1
    Next i
    mOptim.TiempoPiezas = mOptim.Fin
End Sub

Private Sub AnalizaCuerpoDoc(bd As SldWorks.Body2, md As SldWorks.ModelDoc2, p As TPieza)
    On Error GoTo errores
    Set p.Cuerpo = bd
    p.NombreCuerpo = bd.Name
    p.Nombre = bd.Name
    p.Ruta = md.GetPathName
    p.Config = md.ConfigurationManager.ActiveConfiguration.Name
    p.Estado = "OK"
    p.iCaraA = -1: p.iCaraB = -1

    Dim cb As Variant
    cb = bd.GetBodyBox
    If IsEmpty(cb) Then p.Estado = "SIN GEOMETRIA": p.Aviso = "GetBodyBox vacio": Exit Sub
    p.BMin = V3(cb(0), cb(1), cb(2))
    p.BMax = V3(cb(3), cb(4), cb(5))
    p.EnEns = True
    Keep.Add bd

    Dim f As SldWorks.Face2, tr As SldWorks.MathTransform
    Set f = bd.GetFirstFace
    Do While Not f Is Nothing
        AnadeCara f, tr, True, p            ' enEns = True: tr no se usa
        Set f = f.GetNextFace
    Loop
    If p.NCaras = 0 Then p.Estado = "SIN GEOMETRIA": p.Aviso = "Ninguna cara plana valida": Exit Sub
    ClasificaTablero p, tr, True
    Exit Sub
errores:
    p.Estado = "ERROR"
    p.Aviso = "Err " & Err.Number & ": " & Err.Description
End Sub

Private Sub AnalizaComponente(c As SldWorks.Component2, p As TPieza)
    On Error GoTo errores
    Set p.Comp = c
    p.Nombre = c.Name2
    p.Ruta = c.GetPathName
    p.Config = c.ReferencedConfiguration
    p.Estado = "OK"
    p.iCaraA = -1: p.iCaraB = -1

    If c.IsSuppressed Then p.Estado = "SUPRIMIDA": Exit Sub
    If c.IsEnvelope Then p.Estado = "ENVOLVENTE": Exit Sub
    If c.Visible = swComponentHidden Then p.Estado = "OCULTA": Exit Sub

    Dim sup As Long
    sup = c.GetSuppression
    If sup = swComponentLightweight Or sup = swComponentFullyLightweight Then
        p.Estado = "LIGERA"
        Ligeras = Ligeras + 1
        Exit Sub
    End If

    Dim md As SldWorks.ModelDoc2
    Set md = c.GetModelDoc2
    If md Is Nothing Then p.Estado = "SIN DOCUMENTO": Exit Sub
    If md.GetType <> swDocPART Then p.Estado = "ENSAMBLAJE": Exit Sub

    ' --- caja del componente en el ensamblaje ---
    Dim cb As Variant
    cb = c.GetBox(False, False)
    If IsEmpty(cb) Then p.Estado = "SIN GEOMETRIA": p.Aviso = "GetBox vacio": Exit Sub
    p.BMin = V3(cb(0), cb(1), cb(2))
    p.BMax = V3(cb(3), cb(4), cb(5))

    ' --- cuerpos solidos ---
    Dim info As Long, bodies As Variant
    bodies = c.GetBodies3(swSolidBody, info)
    If IsEmpty(bodies) Then
        ' alternativa: cuerpo unico del componente
        Dim b1 As SldWorks.Body2
        Set b1 = c.GetBody
        If Not b1 Is Nothing Then bodies = Array(b1)
    End If
    If IsEmpty(bodies) Then p.Estado = "SIN GEOMETRIA": p.Aviso = "GetBodies3 y GetBody vacios (info=" & info & ")": Exit Sub

    Dim tr As SldWorks.MathTransform
    Set tr = c.GetTotalTransform(False)
    If tr Is Nothing Then Set tr = c.Transform2

    Dim enEns As Boolean, bd As SldWorks.Body2
    Set bd = bodies(0)
    ' Pruebas con SW 2024: GetBodies3 devuelve la geometria en el espacio de la PIEZA
    ' (aunque GetBodyBox venga en el del ensamblaje). Solo si la transformacion es la
    ' identidad coinciden ambos espacios.
    enEns = TransformIdentidad(tr)
    p.EnEns = enEns
    Dim ad As Variant, k3 As Long
    ad = tr.ArrayData
    For k3 = 0 To UBound(ad)
        p.TrDatos = p.TrDatos & Format(ad(k3), "0.000") & " "
    Next k3

    Dim b As Long, f As SldWorks.Face2
    For b = 0 To UBound(bodies)
        Set bd = bodies(b)
        If Not enEns Then
            ' Se transforma una COPIA del cuerpo con la API (maneja bien simetrias
            ' y rotaciones): a partir de aqui todo esta en coordenadas de ensamblaje.
            Set bd = bd.Copy
            bd.ApplyTransform tr
        End If
        Keep.Add bd
        Set f = bd.GetFirstFace
        Do While Not f Is Nothing
            AnadeCara f, tr, True, p
            Set f = f.GetNextFace
        Loop
    Next b

    If p.NCaras = 0 Then p.Estado = "SIN GEOMETRIA": p.Aviso = "Ninguna cara plana valida (cuerpos=" & (UBound(bodies) + 1) & ")": Exit Sub
    ClasificaTablero p, tr, True
    Exit Sub
errores:
    p.Estado = "ERROR"
    p.Aviso = "Err " & Err.Number & ": " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Decide si las coordenadas de los cuerpos devueltos por GetBodies3 ya estan en
' el espacio del ensamblaje o en el de la pieza. Se compara la caja del cuerpo
' (tal cual y transformada) con la caja del componente.
'------------------------------------------------------------------------------
Private Function CuerpoEnEnsamblaje(body As SldWorks.Body2, cb As Variant, _
                                    tr As SldWorks.MathTransform) As Boolean
    Dim bb As Variant
    bb = body.GetBodyBox
    If IsEmpty(bb) Then CuerpoEnEnsamblaje = True: Exit Function

    Dim d1 As Double, k As Long
    For k = 0 To 5
        d1 = d1 + Abs(CDbl(bb(k)) - CDbl(cb(k)))
    Next k

    ' caja transformada (AABB de las 8 esquinas)
    Dim mn As TV3, mx As TV3, i As Long, q As TV3, t As TV3, first As Boolean
    first = True
    For i = 0 To 7
        q = V3(IIf(i And 1, bb(3), bb(0)), IIf(i And 2, bb(4), bb(1)), IIf(i And 4, bb(5), bb(2)))
        t = TrP(tr, q)
        If first Then
            mn = t: mx = t: first = False
        Else
            mn = V3(MinD(mn.X, t.X), MinD(mn.Y, t.Y), MinD(mn.Z, t.Z))
            mx = V3(MaxD(mx.X, t.X), MaxD(mx.Y, t.Y), MaxD(mx.Z, t.Z))
        End If
    Next i
    Dim d2 As Double
    d2 = Abs(mn.X - cb(0)) + Abs(mn.Y - cb(1)) + Abs(mn.Z - cb(2)) + _
         Abs(mx.X - cb(3)) + Abs(mx.Y - cb(4)) + Abs(mx.Z - cb(5))
    CuerpoEnEnsamblaje = (d1 <= d2)
End Function

'------------------------------------------------------------------------------
' Anade una cara plana (si cumple el area minima) a la pieza, en coord. ensamblaje
'------------------------------------------------------------------------------
Private Sub AnadeCara(f As SldWorks.Face2, tr As SldWorks.MathTransform, _
                      enEns As Boolean, p As TPieza)
    Dim s As SldWorks.Surface
    Set s = f.GetSurface
    If s Is Nothing Then Exit Sub
    ' Surface.IsPlane devolvia un valor inconsistente al probarlo en SW 2024;
    ' se usa Identity (4001 = swSurfaceType_e.PLANE_TYPE)
    If s.Identity <> 4001 Then Exit Sub   ' 4001 = superficie plana (swSurfaceType_e)

    Dim a As Double
    a = f.GetArea
    If a < Cfg.AreaMinCara * MM * MM Then Exit Sub

    Dim n As TV3, pt As TV3, pp As Variant, nn As Variant, bx As Variant
    nn = f.Normal
    pp = s.PlaneParams                      ' normal(0..2) + punto(3..5)
    n = VFromArr(nn)
    pt = VFromArr(pp, 3)
    bx = f.GetBox

    If Not enEns Then
        n = TrV(tr, n)
        pt = TrP(tr, pt)
    End If

    Dim c As TCara
    c.N = VUnit(n)
    c.P = pt
    c.Area = a
    Set c.Ref = f

    ' caja alineada a ejes en el ensamblaje
    If enEns Then
        c.BMin = V3(bx(0), bx(1), bx(2))
        c.BMax = V3(bx(3), bx(4), bx(5))
    Else
        Dim i As Long, q As TV3, t As TV3, first As Boolean
        first = True
        For i = 0 To 7
            q = V3(IIf(i And 1, bx(3), bx(0)), IIf(i And 2, bx(4), bx(1)), IIf(i And 4, bx(5), bx(2)))
            t = TrP(tr, q)
            If first Then
                c.BMin = t: c.BMax = t: first = False
            Else
                c.BMin = V3(MinD(c.BMin.X, t.X), MinD(c.BMin.Y, t.Y), MinD(c.BMin.Z, t.Z))
                c.BMax = V3(MaxD(c.BMax.X, t.X), MaxD(c.BMax.Y, t.Y), MaxD(c.BMax.Z, t.Z))
            End If
        Next i
    End If

    ReDim Preserve p.Caras(0 To p.NCaras)
    p.Caras(p.NCaras) = c
    p.NCaras = p.NCaras + 1
End Sub

'------------------------------------------------------------------------------
' Un tablero = cara plana mas grande + otra cara plana paralela y opuesta.
' Espesor = distancia entre ambos planos. Largo/ancho = rectangulo envolvente de
' la cara grande alineado con su arista recta mas larga.
'------------------------------------------------------------------------------
Private Sub ClasificaTablero(p As TPieza, tr As SldWorks.MathTransform, enEns As Boolean)
    Dim i As Long, ia As Long, amax As Double
    ia = 0
    For i = 0 To p.NCaras - 1
        If p.Caras(i).Area > amax Then amax = p.Caras(i).Area: ia = i
    Next i

    Dim ib As Long, bmax As Double
    ib = -1
    For i = 0 To p.NCaras - 1
        If i <> ia Then
            If VDot(p.Caras(i).N, p.Caras(ia).N) < -0.9999 Then
                If p.Caras(i).Area > bmax Then bmax = p.Caras(i).Area: ib = i
            End If
        End If
    Next i
    If ib < 0 Then
        p.Aviso = "No hay cara opuesta paralela: no es un tablero reconocible"
        Dim q As Long
        For q = 0 To p.NCaras - 1
            p.Aviso = p.Aviso & " | N" & q & "=(" & Format(p.Caras(q).N.X, "0.00") & "," & Format(p.Caras(q).N.Y, "0.00") & "," & Format(p.Caras(q).N.Z, "0.00") & ") A=" & Format(p.Caras(q).Area * 1000000#, "0")
        Next q
        Exit Sub
    End If

    Dim d As TV3
    d = VSub(p.Caras(ib).P, p.Caras(ia).P)
    p.Espesor = Abs(VDot(d, p.Caras(ia).N))
    p.iCaraA = ia
    p.iCaraB = ib

    DimensionesCara p, tr, enEns
    If p.Largo > 0 Then
        ' tablero si el espesor es claramente menor que el ancho
        p.EsTablero = (p.Espesor <= 0.4 * p.Ancho)
        If Not p.EsTablero Then p.Aviso = "Espesor demasiado grande respecto al ancho"
    Else
        p.Aviso = "No se pudieron medir largo y ancho"
    End If
End Sub

Private Sub DimensionesCara(p As TPieza, tr As SldWorks.MathTransform, enEns As Boolean)
    Dim f As SldWorks.Face2
    Set f = p.Caras(p.iCaraA).Ref
    Dim edgs As Variant
    edgs = f.GetEdges
    If IsEmpty(edgs) Then Exit Sub

    Dim n As Long, k As Long, e As SldWorks.Edge, v1 As SldWorks.Vertex, v2 As SldWorks.Vertex
    Dim pts() As TV3, np As Long
    Dim uBest As TV3, lBest As Double, a As TV3, b As TV3, l As Double
    n = UBound(edgs)
    ReDim pts(0 To 2 * (n + 1))
    For k = 0 To n
        Set e = edgs(k)
        If Not CBool(e.GetCurve.IsLine) Then p.ContornoCurvo = True
        Set v1 = e.GetStartVertex
        Set v2 = e.GetEndVertex
        If Not v1 Is Nothing And Not v2 Is Nothing Then
            a = VFromArr(v1.GetPoint)
            b = VFromArr(v2.GetPoint)
            If Not enEns Then a = TrP(tr, a): b = TrP(tr, b)
            pts(np) = a: np = np + 1
            pts(np) = b: np = np + 1
            If CBool(e.GetCurve.IsLine) Then
                l = VLen(VSub(b, a))
                If l > lBest Then lBest = l: uBest = VUnit(VSub(b, a))
            End If
        Else
            p.ContornoCurvo = True
        End If
    Next k
    If lBest <= 0 Or np = 0 Then Exit Sub

    Dim nA As TV3, vv As TV3
    nA = p.Caras(p.iCaraA).N
    vv = VUnit(VCross(nA, uBest))
    Dim umin As Double, umax As Double, vmin As Double, vmax As Double, s As Double, t As Double
    umin = 1E+30: vmin = 1E+30: umax = -1E+30: vmax = -1E+30
    For k = 0 To np - 1
        s = VDot(pts(k), uBest)
        t = VDot(pts(k), vv)
        If s < umin Then umin = s
        If s > umax Then umax = s
        If t < vmin Then vmin = t
        If t > vmax Then vmax = t
    Next k
    p.Largo = MaxD(umax - umin, vmax - vmin)
    p.Ancho = MinD(umax - umin, vmax - vmin)
End Sub

'------------------------------------------------------------------------------
' Transformaciones con la API (MathUtility)
'------------------------------------------------------------------------------
Public Function TrP(tr As SldWorks.MathTransform, v As TV3) As TV3
    ' IMPORTANTE: CreatePoint necesita un Double() tipado. Con Array(...) (Variant)
    ' SolidWorks lee mal la memoria y devuelve puntos desplazados (causaba resultados
    ' absurdos en las transformaciones).
    Dim pt As SldWorks.MathPoint, r As Variant, a(0 To 2) As Double
    a(0) = v.X: a(1) = v.Y: a(2) = v.Z
    Set pt = mu.CreatePoint(a)
    Set pt = pt.MultiplyTransform(tr)
    r = pt.ArrayData
    TrP = V3(r(0), r(1), r(2))
End Function

Public Function TrV(tr As SldWorks.MathTransform, v As TV3) As TV3
    ' Vector = diferencia de dos puntos transformados (la traslacion se cancela).
    ' Se evita MathVector.MultiplyTransform: con componentes simetricos daba
    ' resultados inconsistentes en las pruebas con SolidWorks 2024.
    Dim a As TV3, b As TV3
    a = TrP(tr, v)
    b = TrP(tr, V3(0, 0, 0))
    TrV = VSub(a, b)
End Function

' True si la transformacion no mueve nada (rotacion = identidad y traslacion = 0)
Private Function TransformIdentidad(tr As SldWorks.MathTransform) As Boolean
    Dim a As Variant, i As Long
    a = tr.ArrayData
    Dim id As Variant
    id = Array(1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)
    For i = 0 To 11
        If Abs(CDbl(a(i)) - CDbl(id(i))) > 0.000000001 Then Exit Function
    Next i
    TransformIdentidad = True
End Function
