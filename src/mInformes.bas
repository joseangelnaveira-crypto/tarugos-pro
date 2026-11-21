Attribute VB_Name = "mInformes"
Option Explicit
'==============================================================================
' mInformes - FASE 8: informe CSV (separador ";", se abre directo en Excel).
' Una fila por tarugo: union, tarugo, estado, piezas, posicion en el ensamblaje (mm),
' diametro y profundidades, y el motivo si la union era dudosa.
'==============================================================================

Public Sub ExportaCSV(ByVal ruta As String, ByVal ensamblaje As String)
    Dim ff As Integer, u As Long, k As Long, p As TV3, est As String
    ff = FreeFile
    On Error GoTo fallo
    Open ruta For Output As #ff
    Print #ff, "Ensamblaje;" & ensamblaje & ";Fecha;" & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #ff, "Union;Tarugo;Estado union;Estado tarugo;Pieza 1 (prof);Pieza 2 (prof);X mm;Y mm;Z mm;" & _
               "Diametro mm;Prof pieza 1 mm;Prof pieza 2 mm;Motivo"
    For u = 0 To NUniones - 1
        For k = 0 To Uniones(u).NT - 1
            p = Uniones(u).Pos(k)
            If Uniones(u).Activo(k) Then est = "ACTIVO" Else est = "DESCARTADO"
            Print #ff, Uniones(u).Id & ";TARUGO_" & Format(k + 1, "000") & ";" & Uniones(u).Estado & ";" & est & ";" & _
                       Piezas(Uniones(u).P1).Nombre & ";" & Piezas(Uniones(u).P2).Nombre & ";" & _
                       N(p.X * 1000) & ";" & N(p.Y * 1000) & ";" & N(p.Z * 1000) & ";" & _
                       N(Cfg.Diametro) & ";" & N(Cfg.Prof1) & ";" & N(Cfg.Prof2) & ";" & Replace(Uniones(u).Motivo, ";", ",")
        Next k
    Next u
    Close #ff
    Exit Sub
fallo:
    On Error Resume Next
    Close #ff
End Sub

Private Function N(ByVal v As Double) As String
    N = Replace(Format(v, "0.000"), ".", ",")      ' coma decimal para Excel en espanol
End Function
