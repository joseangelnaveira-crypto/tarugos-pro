Attribute VB_Name = "mDuplicados"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mDuplicados - FASE 5: reconocer los taladros creados por esta macro.
'
' Cada operacion de corte creada lleva una propiedad personalizada TARUGO_META:
'     "px;py;pz;nx;ny;nz;D;P"   (mm, en coordenadas de la PIEZA)
' Al volver a ejecutar, se leen las operaciones existentes y no se duplican
' taladros iguales. Las operaciones que NO tengan esa propiedad son ajenas a la
' macro y no se tocan nunca.
'==============================================================================

Public UltimoAdd As String

Public Type TMeta
    Nombre As String
    P As TV3           ' m
    N As TV3
    D As Double        ' mm
    Prof As Double     ' mm
End Type

Public Function Codifica(pp As TV3, nn As TV3, ByVal diam As Double, ByVal prof As Double) As String
    Codifica = F(pp.X * 1000) & ";" & F(pp.Y * 1000) & ";" & F(pp.Z * 1000) & ";" & _
               F(nn.X) & ";" & F(nn.Y) & ";" & F(nn.Z) & ";" & F(diam) & ";" & F(prof)
End Function

Private Function F(ByVal x As Double) As String
    F = Replace(Format(x, "0.000"), ",", ".")      ' independiente del idioma de Windows
End Function

Private Function Num(ByVal s As String) As Double
    Num = Val(Replace(s, ",", "."))               ' Val siempre usa el punto decimal
End Function

' Marca la operacion como creada por la macro
Public Sub Marca(feat As SldWorks.Feature, pp As TV3, nn As TV3, ByVal diam As Double, ByVal prof As Double)
    ' Feature.CustomPropertyManager.Add3 devolvio GenericFail en SW 2024; se guarda en el
    ' COMENTARIO de la operacion (Feature.Description), que viaja con el archivo de la pieza.
    On Error Resume Next
    feat.Description = META_PROP & "=" & Codifica(pp, nn, diam, prof)
    UltimoAdd = "desc=" & feat.Description & " err=" & Err.Number
End Sub

Private Function LeeMeta(ft As SldWorks.Feature) As String
    Dim d As String
    On Error Resume Next
    d = ft.Description
    On Error GoTo 0
    If Left(d, Len(META_PROP) + 1) = META_PROP & "=" Then LeeMeta = Mid(d, Len(META_PROP) + 2)
End Function

' Diagnostico: operaciones de la pieza con su tipo y el valor de TARUGO_META
Public Function Diag(pmd As SldWorks.ModelDoc2) As String
    Dim ft As SldWorks.Feature, v As String, r As String
    Set ft = pmd.FirstFeature
    Do While Not ft Is Nothing
        v = LeeMeta(ft)
        If Left(ft.Name, 6) = "UNION_" Or ft.GetTypeName2 = "ProfileFeature" Then
            r = r & "      [" & ft.GetTypeName2 & "] " & ft.Name & "  meta=" & v & vbCrLf
        End If
        Set ft = ft.GetNextFeature
    Loop
    Diag = r
End Function

' Lee TODAS las operaciones de la macro que ya existen en la pieza
Public Function LeeExistentes(pmd As SldWorks.ModelDoc2, ByRef n As Long) As TMeta()
    Dim r() As TMeta
    n = 0
    Dim ft As SldWorks.Feature, v As String, t() As String
    Set ft = pmd.FirstFeature
    Do While Not ft Is Nothing
        ' el tipo de una operacion de corte aparece como "ICE" en SW 2024 (no "Cut"):
        ' se reconocen por el prefijo del nombre y se confirman con la propiedad TARUGO_META
        If Left(ft.Name, 6) = "UNION_" Then
            v = LeeMeta(ft)
            If Len(v) > 0 Then
                t = Split(v, ";")
                If UBound(t) >= 7 Then
                    ReDim Preserve r(0 To n)
                    r(n).Nombre = ft.Name
                    r(n).P = V3(Num(t(0)) / 1000, Num(t(1)) / 1000, Num(t(2)) / 1000)
                    r(n).N = V3(Num(t(3)), Num(t(4)), Num(t(5)))
                    r(n).D = Num(t(6))
                    r(n).Prof = Num(t(7))
                    n = n + 1
                End If
            End If
        End If
        Set ft = ft.GetNextFeature
    Loop
    LeeExistentes = r
End Function

' 0 = no existe | 1 = ya existe igual | 2 = existe en el mismo sitio pero con otra profundidad/diametro
Public Function EstadoTaladro(ex() As TMeta, ByVal n As Long, pp As TV3, nn As TV3, ByVal diam As Double, ByVal prof As Double) As Long
    Dim i As Long, d As TV3
    For i = 0 To n - 1
        If VDot(ex(i).N, nn) > 0.99 Then
            d = VSub(ex(i).P, pp)
            If VLen(d) < 0.0002 Then                 ' 0,2 mm
                If Abs(ex(i).D - diam) < 0.01 And Abs(ex(i).Prof - prof) < 0.01 Then
                    EstadoTaladro = 1
                Else
                    EstadoTaladro = 2
                End If
                Exit Function
            End If
        End If
    Next i
End Function

' Nombre de la operacion de la macro que ya existe en ese punto/normal ("" si no hay)
Public Function NombreExistente(ex() As TMeta, ByVal n As Long, pp As TV3, nn As TV3) As String
    Dim i As Long
    For i = 0 To n - 1
        If VDot(ex(i).N, nn) > 0.99 Then
            If VLen(VSub(ex(i).P, pp)) < 0.0002 Then NombreExistente = ex(i).Nombre: Exit Function
        End If
    Next i
End Function
