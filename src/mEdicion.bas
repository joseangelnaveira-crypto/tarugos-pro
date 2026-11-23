Attribute VB_Name = "mEdicion"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mEdicion - FASE 3: eliminar / restaurar tarugos y uniones completas.
' Cada tarugo es UN solo elemento (sus dos taladros), asi que nunca puede
' quedar un taladro aislado en una pieza.
'==============================================================================

' Las uniones dudosas empiezan DESACTIVADAS: hay que revisarlas y restaurarlas.
Public Sub EstadoInicial()
    Dim u As Long, k As Long
    For u = 0 To NUniones - 1
        For k = 0 To Uniones(u).NT - 1
            Uniones(u).Activo(k) = (Uniones(u).Estado = "VALIDA") And Not Uniones(u).SinMat(k)
        Next k
    Next u
End Sub

Public Sub EliminaUnion(ByVal u As Long)
    Dim k As Long
    For k = 0 To Uniones(u).NT - 1: Uniones(u).Activo(k) = False: Next k
End Sub

Public Sub RestauraUnion(ByVal u As Long)
    Dim k As Long
    ' los tarugos sin material (hueco) no se reactivan con "restaurar union" / "aceptar todos"
    For k = 0 To Uniones(u).NT - 1
        If Not Uniones(u).SinMat(k) Then Uniones(u).Activo(k) = True
    Next k
End Sub

' Reactiva todos los tarugos de las uniones VALIDAS (las dudosas no se tocan)
Public Sub AceptarTodos()
    Dim u As Long
    For u = 0 To NUniones - 1
        If Uniones(u).Estado = "VALIDA" Then RestauraUnion u
    Next u
End Sub

Public Function ContarActivos() As Long
    Dim u As Long, k As Long, n As Long
    For u = 0 To NUniones - 1
        For k = 0 To Uniones(u).NT - 1
            If Uniones(u).Activo(k) Then n = n + 1
        Next k
    Next u
    ContarActivos = n
End Function

Public Function ContarTarugos() As Long
    Dim u As Long, n As Long
    For u = 0 To NUniones - 1: n = n + Uniones(u).NT: Next u
    ContarTarugos = n
End Function

' Mantiene solo las uniones cuyas DOS piezas estan en la lista de nombres (modo manual)
Public Sub FiltraPorPiezas(nombres As Collection)
    Dim keep() As TUnion, nk As Long, u As Long
    For u = 0 To NUniones - 1
        If EstaEn(nombres, Piezas(Uniones(u).P1).Nombre) And EstaEn(nombres, Piezas(Uniones(u).P2).Nombre) Then
            ReDim Preserve keep(0 To nk)
            keep(nk) = Uniones(u)
            nk = nk + 1
        End If
    Next u
    Erase Uniones
    NUniones = nk
    If nk > 0 Then
        ReDim Uniones(0 To nk - 1)
        For u = 0 To nk - 1: Uniones(u) = keep(u): Next u
    End If
End Sub

Private Function EstaEn(c As Collection, ByVal s As String) As Boolean
    Dim v As Variant
    For Each v In c
        ' una pieza cuenta si es la seleccionada o esta DENTRO de un subensamblaje seleccionado
        ' (Name2 de una pieza anidada = "sub-1/pieza-1")
        If CStr(v) = s Then EstaEn = True: Exit Function
        If Left(s, Len(CStr(v)) + 1) = CStr(v) & "/" Then EstaEn = True: Exit Function
    Next v
End Function
