Attribute VB_Name = "mVec"
Option Explicit
'==============================================================================
' mVec - Tipo vector 3D y operaciones basicas (sin llamadas a SolidWorks).
'==============================================================================

Public Type TV3
    X As Double
    Y As Double
    Z As Double
End Type

Public Function V3(ByVal x As Double, ByVal y As Double, ByVal z As Double) As TV3
    V3.X = x: V3.Y = y: V3.Z = z
End Function

Public Function VDot(a As TV3, b As TV3) As Double
    VDot = a.X * b.X + a.Y * b.Y + a.Z * b.Z
End Function

Public Function VCross(a As TV3, b As TV3) As TV3
    VCross.X = a.Y * b.Z - a.Z * b.Y
    VCross.Y = a.Z * b.X - a.X * b.Z
    VCross.Z = a.X * b.Y - a.Y * b.X
End Function

Public Function VSub(a As TV3, b As TV3) As TV3
    VSub.X = a.X - b.X: VSub.Y = a.Y - b.Y: VSub.Z = a.Z - b.Z
End Function

Public Function VLen(a As TV3) As Double
    VLen = Sqr(a.X * a.X + a.Y * a.Y + a.Z * a.Z)
End Function

Public Function VUnit(a As TV3) As TV3
    Dim l As Double
    l = VLen(a)
    If l > 0 Then
        VUnit.X = a.X / l: VUnit.Y = a.Y / l: VUnit.Z = a.Z / l
    End If
End Function

Public Function VFromArr(arr As Variant, Optional ByVal i0 As Long = 0) As TV3
    VFromArr.X = CDbl(arr(i0))
    VFromArr.Y = CDbl(arr(i0 + 1))
    VFromArr.Z = CDbl(arr(i0 + 2))
End Function

Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function
