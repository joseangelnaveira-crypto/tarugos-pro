Attribute VB_Name = "mMain"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
' Ensamblaje que se analizo en la ultima revision (ruta|titulo). mTaladros lo compara con el
' documento activo: la ventana de revision no es modal y, si se cambia de documento, Piezas()
' y Uniones() serian de otro ensamblaje.
Public AsmAnalizado As String
'==============================================================================
' FASE 3: previsualizacion 3D y edicion de tarugos (NO modifica las piezas).
' AUTOMATICO: sin nada seleccionado analiza todo el ensamblaje.
' MANUAL    : con 2 o mas componentes seleccionados, solo las uniones entre ellos.
' Ejecutar el procedimiento  Fase3_Revisar.
'==============================================================================
Public Sub Fase3_Revisar(Optional ByVal manual As Boolean = False)
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    Dim md As SldWorks.ModelDoc2
    Set md = swApp.ActiveDoc
    If md Is Nothing Then MsgBox "Abre un ensamblaje.", vbExclamation, "Tarugos": Exit Sub
    If md.GetType <> swDocASSEMBLY And md.GetType <> swDocPART Then MsgBox "El documento activo debe ser un ensamblaje o una pieza multicuerpo.", vbExclamation, "Tarugos": Exit Sub
    If md.GetType = swDocPART Then
        If NCuerposSolidos(md) < 2 Then MsgBox "La pieza activa tiene un solo cuerpo solido: abre un ensamblaje o una pieza multicuerpo.", vbExclamation, "Tarugos": Exit Sub
    End If

    ' si ya hay una revision abierta, se cierra
    mPreview.Limpia

    CfgPorDefecto
    ' --- modo manual: piezas que el usuario ha seleccionado con el raton ---
    ' (vale seleccionar la pieza en el arbol o pinchar cualquier cara/arista de la pieza)
    Dim sm As SldWorks.SelectionMgr
    Set sm = md.SelectionManager
    Dim nombres As Collection
    Set nombres = New Collection
    Dim i As Long, n As Long, c As SldWorks.Component2
    If manual Then
        n = sm.GetSelectedObjectCount2(-1)
        For i = 1 To n
            Set c = Nothing
            If md.GetType = swDocPART Then
                ' pieza multicuerpo: se seleccionan CUERPOS (en el arbol o pinchando en el modelo)
                Dim cuerpoSel As SldWorks.Body2
                Set cuerpoSel = CuerpoDeSeleccion(sm, i)
                If Not cuerpoSel Is Nothing Then
                    If Not YaEsta(nombres, cuerpoSel.Name) Then nombres.Add cuerpoSel.Name
                End If
            ElseIf sm.GetSelectedObjectType3(i, -1) = swSelCOMPONENTS Then
                Set c = sm.GetSelectedObject6(i, -1)
            Else
                Set c = sm.GetSelectedObjectsComponent4(i, -1)
            End If
            If Not c Is Nothing Then
                If Not YaEsta(nombres, c.Name2) Then nombres.Add c.Name2
            End If
        Next i
        If nombres.Count < 2 Then
            MsgBox "Modo MANUAL: selecciona con el raton 2 o mas piezas del ensamblaje (Ctrl + clic para varias) y pulsa CONTINUAR.", _
                   vbInformation, "Tarugos"
            Exit Sub
        End If
    End If

    If md.GetType = swDocPART Then
        AnalizaPiezasDoc swApp, md          ' pieza multicuerpo: cada cuerpo solido es una pieza
    Else
        Dim asm As SldWorks.AssemblyDoc
        Set asm = md
        AnalizaPiezas swApp, asm
    End If
    If mOptim.Cancelado Then MsgBox "Analisis cancelado. No se ha modificado nada.", vbInformation, "Tarugos": Exit Sub
    DetectaUniones
    If mOptim.Cancelado Then MsgBox "Analisis cancelado. No se ha modificado nada.", vbInformation, "Tarugos": Exit Sub
    Dim modo As String
    modo = "AUTOMATICO"
    If manual Then
        modo = "MANUAL"
        FiltraPorPiezas nombres
    End If
    EscribeUniones md
    If NUniones = 0 Then
        MsgBox "No se han encontrado uniones (modo " & modo & ").", vbInformation, "Tarugos"
        Exit Sub
    End If

    EstadoInicial
    AsmAnalizado = md.GetPathName & "|" & md.GetTitle
    md.ClearSelection2 True
    ConstruyePreview swApp, md

    Set Eventos = New cEventos
    Eventos.Conecta md
    Set Formulario = frmRevision
    frmRevision.Caption = "Tarugos - Revision (" & modo & ")"
    frmRevision.Refresca
    frmRevision.Show vbModeless
    DibujaTodo
    Redibuja
End Sub


'==============================================================================
' PUNTO DE ENTRADA: ejecutar  Tarugos_Principal
' Ventana con la configuracion y los botones AUTOMATICO / MANUAL / DESHACER.
' Para MANUAL: selecciona antes (2 o mas) componentes en el ensamblaje.
'==============================================================================
Public Sub Tarugos_Principal()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    If swApp.ActiveDoc Is Nothing Then MsgBox "Abre un ensamblaje.", vbExclamation, "Tarugos": Exit Sub
    If swApp.ActiveDoc.GetType <> swDocASSEMBLY And swApp.ActiveDoc.GetType <> swDocPART Then MsgBox "El documento activo debe ser un ensamblaje o una pieza multicuerpo.", vbExclamation, "Tarugos": Exit Sub
    CfgPorDefecto
    ' Ventana NO modal: asi se puede seleccionar piezas con el raton mientras esta abierta.
    ' Los botones llaman directamente a Fase3_Revisar / Fase6_Deshacer.
    frmPrincipal.Reinicia
    frmPrincipal.Show vbModeless
End Sub

' Numero de cuerpos solidos visibles de una pieza
Private Function NCuerposSolidos(md As SldWorks.ModelDoc2) As Long
    Dim pd As SldWorks.PartDoc, b As Variant
    Set pd = md
    b = pd.GetBodies2(swSolidBody, True)
    If Not IsEmpty(b) Then NCuerposSolidos = UBound(b) + 1
End Function

Private Function YaEsta(c As Collection, ByVal s As String) As Boolean
    Dim v As Variant
    For Each v In c
        If CStr(v) = s Then YaEsta = True: Exit Function
    Next v
End Function

' Cuerpo solido al que pertenece lo seleccionado en una PIEZA: el propio cuerpo (arbol)
' o el cuerpo de la cara pinchada en el modelo.
Private Function CuerpoDeSeleccion(sm As SldWorks.SelectionMgr, ByVal i As Long) As SldWorks.Body2
    Dim t As Long, o As Object, f As SldWorks.Face2
    On Error Resume Next
    t = sm.GetSelectedObjectType3(i, -1)
    Set o = sm.GetSelectedObject6(i, -1)
    If o Is Nothing Then Exit Function
    If t = swSelSOLIDBODIES Then
        Set CuerpoDeSeleccion = o
    ElseIf t = swSelFACES Then
        Set f = o
        Set CuerpoDeSeleccion = f.GetBody
    End If
End Function

' Lista TODAS las uniones (tambien las dudosas o descartadas, que pueden no tener tarugos)
' con su motivo, en Tarugos_Uniones.txt junto al documento.
Private Sub EscribeUniones(md As SldWorks.ModelDoc2)
    Dim carpeta As String, ruta As String, ff As Integer, i As Long
    carpeta = Environ("TEMP")
    If md.GetPathName <> "" Then carpeta = Left(md.GetPathName, InStrRev(md.GetPathName, "\") - 1)
    ruta = carpeta & "\Tarugos_Uniones.txt"
    On Error GoTo fallo
    ff = FreeFile
    Open ruta For Output As #ff
    Print #ff, "TARUGOS PRO - uniones detectadas   " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "   " & md.GetTitle
    For i = 0 To NPiezas - 1
        Print #ff, "PIEZA [" & i & "] " & Piezas(i).Nombre & "  " & Piezas(i).Estado & _
                   IIf(Piezas(i).EsTablero, "  tablero, espesor " & Format(Piezas(i).Espesor * 1000, "0.0") & " mm", "") & _
                   IIf(Piezas(i).Aviso <> "", "  AVISO: " & Piezas(i).Aviso, "")
    Next i
    For i = 0 To NUniones - 1
        With Uniones(i)
            Print #ff, .Id & "  " & .Estado & "  " & .Tipo & "  tarugos=" & .NT & "  L=" & Format(.L * 1000, "0.0") & _
                       " W=" & Format(.W * 1000, "0.0") & "  " & Piezas(.P1).Nombre & " / " & Piezas(.P2).Nombre
            If .Motivo <> "" Then Print #ff, "      MOTIVO: " & .Motivo
        End With
    Next i
    Close #ff
    Exit Sub
fallo:
    On Error Resume Next
    Close #ff
End Sub
