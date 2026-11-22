Attribute VB_Name = "mRecuperacion"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mRecuperacion - FASE 6: deshacer el mecanizado.
'
' Fase6_Deshacer elimina SOLO las operaciones creadas por esta macro (las que
' llevan la marca TARUGO_META en su comentario), con todos sus croquis, en todas
' las piezas del ensamblaje activo. Nunca toca otras operaciones de diseno.
' Antes de borrar hace una copia de seguridad de cada pieza y pide confirmacion.
' Piezas virtuales, sin archivo o de solo lectura se omiten y se anotan.
'==============================================================================

Public Sub Fase6_Deshacer()
    Dim app As SldWorks.SldWorks
    Set app = Application.SldWorks
    Dim amd As SldWorks.ModelDoc2
    Set amd = app.ActiveDoc
    If amd Is Nothing Then Exit Sub
    If amd.GetType = swDocPART Then DeshacerPieza app, amd: Exit Sub     ' pieza multicuerpo
    If amd.GetType <> swDocASSEMBLY Then MsgBox "Activa el ensamblaje.", vbExclamation, "Tarugos": Exit Sub
    Dim asm As SldWorks.AssemblyDoc
    Set asm = amd
    Dim titAsm As String
    titAsm = amd.GetTitle

    Dim inf As String
    Dim comps As Variant
    comps = asm.GetComponents(False)
    If IsEmpty(comps) Then Exit Sub

    ' ---- documentos distintos (una pieza repetida se trata una sola vez) ----
    Dim rutas() As String, nR As Long, i As Long, j As Long, c As SldWorks.Component2, md As SldWorks.ModelDoc2, ya As Boolean
    ReDim rutas(0 To UBound(comps))
    For i = 0 To UBound(comps)
        Set c = comps(i)
        If c.IsSuppressed = 0 Then
            Set md = c.GetModelDoc2
            If md Is Nothing Then
                ' ligero o sin cargar: se ignora
            ElseIf md.GetType <> swDocPART Then
                ' subensamblaje: sus piezas ya vienen en la lista
            ElseIf c.IsVirtual <> 0 Then
                inf = inf & "OMITIDA " & c.Name2 & ": pieza virtual (sin archivo propio)" & vbCrLf
            ElseIf md.GetPathName = "" Then
                inf = inf & "OMITIDA " & c.Name2 & ": sin archivo" & vbCrLf
            ElseIf mGuardado.EsSoloLectura(md) Then
                inf = inf & "OMITIDA " & c.Name2 & ": solo lectura" & vbCrLf
            Else
                ya = False
                For j = 0 To nR - 1
                    If LCase(rutas(j)) = LCase(md.GetPathName) Then ya = True: Exit For
                Next j
                If Not ya Then rutas(nR) = md.GetPathName: nR = nR + 1
            End If
        End If
    Next i

    ' ---- contar operaciones de la macro por documento ----
    Dim cuenta() As Long, total As Long, ex() As mDuplicados.TMeta, nex As Long, resumen As String
    ReDim cuenta(0 To nR)
    For i = 0 To nR - 1
        Set md = app.GetOpenDocumentByName(rutas(i))
        If md Is Nothing Then
            Dim e As Long, w As Long
            app.DocumentVisible False, swDocPART
            Set md = app.OpenDoc6(rutas(i), swDocPART, swOpenDocOptions_Silent, "", e, w)
            app.DocumentVisible True, swDocPART
        End If
        If Not md Is Nothing Then
            ex = mDuplicados.LeeExistentes(md, nex)
            cuenta(i) = nex
            total = total + nex
            If nex > 0 Then resumen = resumen & "  " & rutas(i) & "  (" & nex & " operaciones)" & vbCrLf
        End If
    Next i
    app.ActivateDoc3 titAsm, False, 1, e

    If total = 0 Then
        MsgBox "No hay operaciones de la macro que deshacer." & vbCrLf & vbCrLf & inf, vbInformation, "Tarugos"
        Exit Sub
    End If
    If MsgBox("Se van a ELIMINAR " & total & " operaciones creadas por la macro (con sus croquis) en:" & vbCrLf & vbCrLf & resumen & vbCrLf & _
              "No se tocan las demas operaciones de las piezas. Antes se hace copia de seguridad." & vbCrLf & "Â¿Continuar?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Tarugos - deshacer") <> vbYes Then Exit Sub

    ' ---- copia de seguridad y borrado ----
    Dim carpeta As String
    carpeta = mGuardado.CarpetaRespaldo(amd)
    Dim borradas As Long, n As Long
    For i = 0 To nR - 1
        If cuenta(i) > 0 Then
            Set md = app.GetOpenDocumentByName(rutas(i))
            If md Is Nothing Then
                app.DocumentVisible False, swDocPART
                Set md = app.OpenDoc6(rutas(i), swDocPART, swOpenDocOptions_Silent, "", e, w)
                app.DocumentVisible True, swDocPART
            End If
            If Not mGuardado.CopiaA(md, carpeta & "\" & Mid(rutas(i), InStrRev(rutas(i), "\") + 1)) Then
                inf = inf & "OMITIDA " & rutas(i) & ": no se pudo hacer la copia de seguridad" & vbCrLf
            Else
                n = BorraOperaciones(md)
                borradas = borradas + n
                If n > 0 Then
                    Dim er As Long
                    er = mGuardado.GuardaDoc(md)
                    If er <> 0 Then
                        inf = inf & "AVISO error al guardar " & rutas(i) & " (codigo " & er & ")" & vbCrLf
                    Else
                        inf = inf & "Guardada " & rutas(i) & ": " & n & " operaciones eliminadas" & vbCrLf
                    End If
                End If
            End If
        End If
    Next i

    app.ActivateDoc3 titAsm, False, 1, e
    Set amd = app.ActiveDoc
    amd.ForceRebuild3 False

    Dim ruta As String, ff As Integer
    ruta = carpeta & "\Informe_Deshacer.txt"
    ff = FreeFile
    Open ruta For Output As #ff
    Print #ff, "TARUGOS PRO - Fase 6 - Deshacer mecanizado   " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #ff, "Operaciones eliminadas: " & borradas & " de " & total
    Print #ff, inf
    Close #ff
    MsgBox "Operaciones eliminadas: " & borradas & " de " & total & vbCrLf & vbCrLf & "Informe: " & ruta, vbInformation, "Tarugos - Fase 6"
End Sub

' Pieza (multicuerpo) activa: mismo proceso sobre su unico documento
Private Sub DeshacerPieza(app As SldWorks.SldWorks, md As SldWorks.ModelDoc2)
    If md.GetPathName = "" Then MsgBox "Guarda la pieza antes de deshacer.", vbExclamation, "Tarugos": Exit Sub
    If mGuardado.EsSoloLectura(md) Then MsgBox "La pieza es de solo lectura.", vbExclamation, "Tarugos": Exit Sub
    Dim ex() As mDuplicados.TMeta, nex As Long
    ex = mDuplicados.LeeExistentes(md, nex)
    If nex = 0 Then MsgBox "No hay operaciones de la macro que deshacer.", vbInformation, "Tarugos": Exit Sub
    If MsgBox("Se van a ELIMINAR " & nex & " operaciones creadas por la macro (con sus croquis) en:" & vbCrLf & vbCrLf & _
              "  " & md.GetPathName & vbCrLf & vbCrLf & _
              "No se tocan las demas operaciones. Antes se hace copia de seguridad." & vbCrLf & "¿Continuar?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Tarugos - deshacer") <> vbYes Then Exit Sub
    Dim carpeta As String, n As Long, er As Long
    carpeta = mGuardado.CarpetaRespaldo(md)
    If Not mGuardado.CopiaA(md, carpeta & "\" & Mid(md.GetPathName, InStrRev(md.GetPathName, "\") + 1)) Then
        MsgBox "No se pudo hacer la copia de seguridad: no se elimina nada.", vbExclamation, "Tarugos": Exit Sub
    End If
    n = BorraOperaciones(md)
    If n > 0 Then er = mGuardado.GuardaDoc(md)
    MsgBox "Operaciones eliminadas: " & n & " de " & nex & IIf(er <> 0, vbCrLf & "AVISO: error al guardar (codigo " & er & ")", "") & _
           vbCrLf & vbCrLf & "Copia de seguridad: " & carpeta, vbInformation, "Tarugos - Fase 6"
End Sub

' Elimina de la pieza las operaciones con marca TARUGO_META. Devuelve cuantas.
Private Function BorraOperaciones(md As SldWorks.ModelDoc2) As Long
    Dim ex() As mDuplicados.TMeta, nex As Long, k As Long, ft As SldWorks.Feature, hechos As Long
    Dim pd As SldWorks.PartDoc
    Set pd = md
    ex = mDuplicados.LeeExistentes(md, nex)
    For k = 0 To nex - 1
        Set ft = pd.FeatureByName(ex(k).Nombre)
        If Not ft Is Nothing Then
            md.ClearSelection2 True
            ft.Select2 False, 0
            If md.Extension.DeleteSelection2(swDelete_Absorbed) Then hechos = hechos + 1
        End If
    Next k
    md.ClearSelection2 True
    BorraOperaciones = hechos
End Function
