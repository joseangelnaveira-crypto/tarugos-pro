Attribute VB_Name = "mMain"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
Private Const DEBUG_INFORME As Boolean = False   ' True: vuelca caras y transformaciones al informe
' Ensamblaje que se analizo en la ultima revision (ruta|titulo). mTaladros lo compara con el
' documento activo: la ventana de revision no es modal y, si se cambia de documento, Piezas()
' y Uniones() serian de otro ensamblaje.
Public AsmAnalizado As String
' TARUGOS PRO - Fases 1 y 2. Ejecutar Fase2_Analizar (uniones y tarugos) o
' Fase1_Analizar (solo piezas) con un ensamblaje abierto. No modifican documentos.
' ---
' ---
' ---
Public Sub Fase2_Analizar()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    Dim md As SldWorks.ModelDoc2
    Set md = swApp.ActiveDoc
    If md Is Nothing Then MsgBox "Abre un ensamblaje.", vbExclamation, "Tarugos": Exit Sub
    If md.GetType <> swDocASSEMBLY Then MsgBox "El documento activo debe ser un ensamblaje.", vbExclamation, "Tarugos": Exit Sub

    CfgPorDefecto
    Dim t0 As Double
    t0 = Timer

    Dim asm As SldWorks.AssemblyDoc
    Set asm = md
    AnalizaPiezas swApp, asm
    If mOptim.Cancelado Then MsgBox "Analisis cancelado. No se ha modificado nada.", vbInformation, "Tarugos": Exit Sub
    DetectaUniones
    If mOptim.Cancelado Then MsgBox "Analisis cancelado. No se ha modificado nada.", vbInformation, "Tarugos": Exit Sub

    Dim seg As Double
    seg = Timer - t0

    Dim ruta As String
    ruta = EscribeInforme2(swApp, md, seg)

    Dim nV As Long, nD As Long, nX As Long, nT As Long, i As Long
    For i = 0 To NUniones - 1
        Select Case Uniones(i).Estado
            Case "VALIDA": nV = nV + 1: nT = nT + Uniones(i).NT
            Case "DUDOSA": nD = nD + 1
            Case Else: nX = nX + 1
        End Select
    Next i

    MsgBox "Uniones validas: " & nV & vbCrLf & _
           "Uniones dudosas: " & nD & vbCrLf & _
           "Descartadas: " & nX & vbCrLf & _
           "Tarugos calculados (validas): " & nT & vbCrLf & _
           "Tiempo: " & Format(seg, "0.00") & " s" & vbCrLf & vbCrLf & _
           "Informe: " & ruta, vbInformation, "Tarugos - Fase 2"
End Sub

Private Function EscribeInforme2(swApp As SldWorks.SldWorks, md As SldWorks.ModelDoc2, seg As Double) As String
    Dim carpeta As String
    carpeta = Environ("TEMP")
    If md.GetPathName <> "" Then carpeta = Left(md.GetPathName, InStrRev(md.GetPathName, "\") - 1)
    Dim ruta As String
    ruta = carpeta & "\Tarugos_Fase2_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim ff As Integer
    ff = FreeFile
    Open ruta For Output As #ff
    Print #ff, "TARUGOS PRO - Fase 2 - Uniones y posiciones"
    Print #ff, "SolidWorks revision: " & swApp.RevisionNumber & "   Ensamblaje: " & md.GetTitle
    Print #ff, "Fecha: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "   Tiempo: " & Format(seg, "0.00") & " s"
    Print #ff, "Rendimiento: piezas " & Format(mOptim.TiempoPiezas, "0.00") & " s, uniones " & Format(mOptim.TiempoUniones, "0.00") & _
               " s; parejas revisadas " & mOptim.ParesRevisados & ", con cajas que se tocan " & mOptim.ParesCandidatos
    Print #ff, "Parametros: diametro " & Cfg.Diametro & " mm, prof1 " & Cfg.Prof1 & " mm, prof2 " & Cfg.Prof2 & _
               " mm, extremos " & Cfg.DistExtremos & " mm"
    Print #ff, String(70, "-")
    Dim i As Long, k As Long
    Print #ff, "PIEZAS"
    For i = 0 To NPiezas - 1
        With Piezas(i)
            Dim q As Long
            Print #ff, "  [" & i & "] " & .Nombre & "  " & .Estado & IIf(.EsTablero, "  tablero " & F1(.Largo) & " x " & F1(.Ancho) & " x " & F1(.Espesor), "")
            If .Estado = "OK" Then Print #ff, "      caja (mm): (" & F1(.BMin.X) & " ; " & F1(.BMin.Y) & " ; " & F1(.BMin.Z) & ") - (" & F1(.BMax.X) & " ; " & F1(.BMax.Y) & " ; " & F1(.BMax.Z) & ")"
            If DEBUG_INFORME Then
            Print #ff, "      enEns=" & .EnEns & "  TR=" & .TrDatos
            For q = 0 To .NCaras - 1
                Print #ff, "      cara" & q & " N=(" & F3(.Caras(q).N.X) & "," & F3(.Caras(q).N.Y) & "," & F3(.Caras(q).N.Z) & ") P=(" & F1(.Caras(q).P.X) & "," & F1(.Caras(q).P.Y) & "," & F1(.Caras(q).P.Z) & ") box=(" & _
                           F1(.Caras(q).BMin.X) & "," & F1(.Caras(q).BMin.Y) & "," & F1(.Caras(q).BMin.Z) & ")-(" & F1(.Caras(q).BMax.X) & "," & F1(.Caras(q).BMax.Y) & "," & F1(.Caras(q).BMax.Z) & ")"
            Next q
            End If
        End With
    Next i
    Print #ff, String(70, "-")
    Print #ff, "UNIONES: " & NUniones
    For i = 0 To NUniones - 1
        With Uniones(i)
            Print #ff, .Id & "  " & .Estado & "  " & .Tipo
            Print #ff, "    Pieza 1 (prof " & Cfg.Prof1 & "): " & Piezas(.P1).Nombre
            Print #ff, "    Pieza 2 (prof " & Cfg.Prof2 & "): " & Piezas(.P2).Nombre
            Print #ff, "    Longitud: " & F1(.L) & " mm   Ancho contacto: " & F1(.W) & " mm   Tarugos: " & .NT
            Print #ff, "    Direccion: (" & F3(.Dir.X) & "," & F3(.Dir.Y) & "," & F3(.Dir.Z) & ")   Normal pieza 1: (" & _
                       F3(.N1.X) & "," & F3(.N1.Y) & "," & F3(.N1.Z) & ")"
            If .Motivo <> "" Then Print #ff, "    MOTIVO: " & .Motivo
            If .NT > 0 Then
                For k = 0 To .NT - 1
                    Print #ff, "    TARUGO_" & Format(k + 1, "000") & "  (" & F1(.Pos(k).X) & " ; " & F1(.Pos(k).Y) & " ; " & F1(.Pos(k).Z) & ") mm"
                Next k
            End If
        End With
    Next i
    Close #ff
    EscribeInforme2 = ruta
End Function

Public Sub Fase1_Analizar()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    Dim md As SldWorks.ModelDoc2
    Set md = swApp.ActiveDoc
    If md Is Nothing Then MsgBox "Abre un ensamblaje.", vbExclamation, "Tarugos": Exit Sub
    If md.GetType <> swDocASSEMBLY Then MsgBox "El documento activo debe ser un ensamblaje.", vbExclamation, "Tarugos": Exit Sub

    CfgPorDefecto
    Dim t0 As Double
    t0 = Timer

    Dim asm As SldWorks.AssemblyDoc
    Set asm = md
    AnalizaPiezas swApp, asm

    Dim seg As Double
    seg = Timer - t0

    Dim ruta As String
    ruta = EscribeInforme(swApp, md, seg)

    Dim nOk As Long, nTab As Long, i As Long
    For i = 0 To NPiezas - 1
        If Piezas(i).Estado = "OK" Then nOk = nOk + 1
        If Piezas(i).EsTablero Then nTab = nTab + 1
    Next i

    MsgBox "Componentes: " & NPiezas & vbCrLf & _
           "Analizados (OK): " & nOk & vbCrLf & _
           "Tableros reconocidos: " & nTab & vbCrLf & _
           "Ligeros sin analizar: " & Ligeras & vbCrLf & _
           "Tiempo: " & Format(seg, "0.00") & " s" & vbCrLf & vbCrLf & _
           "Informe: " & ruta, vbInformation, "Tarugos - Fase 1"
End Sub

Private Function EscribeInforme(swApp As SldWorks.SldWorks, md As SldWorks.ModelDoc2, seg As Double) As String
    Dim carpeta As String
    carpeta = Environ("TEMP")
    If md.GetPathName <> "" Then carpeta = Left(md.GetPathName, InStrRev(md.GetPathName, "\") - 1)

    Dim ruta As String
    ruta = carpeta & "\Tarugos_Fase1_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim ff As Integer
    ff = FreeFile
    Open ruta For Output As #ff
    Print #ff, "TARUGOS PRO - Fase 1 - Reconocimiento de piezas"
    Print #ff, "SolidWorks revision: " & swApp.RevisionNumber
    Print #ff, "Ensamblaje: " & md.GetTitle
    Print #ff, "Fecha: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #ff, "Tiempo de analisis: " & Format(seg, "0.00") & " s"
    Print #ff, "Componentes: " & NPiezas
    Print #ff, String(70, "-")

    Dim i As Long, j As Long
    For i = 0 To NPiezas - 1
        With Piezas(i)
            Print #ff, "[" & i & "] " & .Nombre & "  (" & .Estado & ")"
            Print #ff, "    Archivo: " & .Ruta & "  Config: " & .Config
            If .Estado = "OK" Then
                Print #ff, "    Caja (mm): (" & F1(.BMin.X) & "," & F1(.BMin.Y) & "," & F1(.BMin.Z) & ") - (" & _
                           F1(.BMax.X) & "," & F1(.BMax.Y) & "," & F1(.BMax.Z) & ")"
                Print #ff, "    Caras planas: " & .NCaras
                Print #ff, "    Tablero: " & IIf(.EsTablero, "SI", "NO") & _
                           "  Espesor: " & F1(.Espesor) & "  Largo: " & F1(.Largo) & "  Ancho: " & F1(.Ancho) & _
                           IIf(.ContornoCurvo, "  (contorno curvo)", "")
                If .iCaraA >= 0 Then
                    Print #ff, "    Cara A normal: (" & F3(.Caras(.iCaraA).N.X) & "," & F3(.Caras(.iCaraA).N.Y) & "," & F3(.Caras(.iCaraA).N.Z) & ")"
                End If
            End If
            If .Aviso <> "" Then Print #ff, "    AVISO: " & .Aviso
        End With
    Next i
    Close #ff
    EscribeInforme = ruta
End Function

Private Function F1(ByVal m As Double) As String
    F1 = Format(m * 1000, "0.0")
End Function

Private Function F3(ByVal v As Double) As String
    F3 = Format(v, "0.000")
End Function


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
    If md.GetType <> swDocASSEMBLY Then MsgBox "El documento activo debe ser un ensamblaje.", vbExclamation, "Tarugos": Exit Sub

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
            If sm.GetSelectedObjectType3(i, -1) = swSelCOMPONENTS Then
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

    Dim asm As SldWorks.AssemblyDoc
    Set asm = md
    AnalizaPiezas swApp, asm
    If mOptim.Cancelado Then MsgBox "Analisis cancelado. No se ha modificado nada.", vbInformation, "Tarugos": Exit Sub
    DetectaUniones
    If mOptim.Cancelado Then MsgBox "Analisis cancelado. No se ha modificado nada.", vbInformation, "Tarugos": Exit Sub
    Dim modo As String
    modo = "AUTOMATICO"
    If manual Then
        modo = "MANUAL"
        FiltraPorPiezas nombres
    End If
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
    If swApp.ActiveDoc.GetType <> swDocASSEMBLY Then MsgBox "El documento activo debe ser un ensamblaje.", vbExclamation, "Tarugos": Exit Sub
    CfgPorDefecto
    ' Ventana NO modal: asi se puede seleccionar piezas con el raton mientras esta abierta.
    ' Los botones llaman directamente a Fase3_Revisar / Fase6_Deshacer.
    frmPrincipal.Reinicia
    frmPrincipal.Show vbModeless
End Sub

Private Function YaEsta(c As Collection, ByVal s As String) As Boolean
    Dim v As Variant
    For Each v In c
        If CStr(v) = s Then YaEsta = True: Exit Function
    Next v
End Function
