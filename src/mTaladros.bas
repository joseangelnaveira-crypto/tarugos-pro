Attribute VB_Name = "mTaladros"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
Private croquisPrevios As Collection          ' nombres de croquis previos (ver SnapshotCroquis)
'==============================================================================
' mTaladros - FASES 4 y 5: taladros REALES en las piezas individuales.
'
' Para cada tarugo ACTIVO se crea, en el documento de cada una de las dos piezas,
' un Corte-extruido circular (editable en el arbol de operaciones) sobre la cara
' de contacto, de diametro Cfg.Diametro y profundidad Prof1 (pieza 1) o Prof2
' (pieza 2). Los dos taladros de un tarugo comparten eje y nombre:
'     UNION_001_TARUGO_002_<pieza>_D8_P31
' Antes de tocar nada: confirmacion, copia de seguridad del archivo de cada pieza,
' y verificacion del volumen eliminado.
' Piezas virtuales (guardadas dentro del ensamblaje) y piezas compartidas por
' varias instancias se OMITEN y se anotan en el informe (se resuelven en la Fase 5).
'==============================================================================

Private Type TJob
    U As Long
    K As Long
    Pieza As Long
    Pc As TV3          ' centro del taladro en el plano de contacto (ensamblaje)
    Nout As TV3        ' normal saliente de la cara de la pieza (ensamblaje)
    Prof As Double     ' m
    Nombre As String
    Borrar As Boolean  ' tarugo DESCARTADO: si ya estaba taladrado, se elimina
End Type

Private Jobs() As TJob
Private NJobs As Long
Private inf As String
Private nSaltados As Long
Private nBorrados As Long
Private hayAviso As Boolean            ' un taladro de la pieza actual elimino un volumen inesperado

' --- tiempos (se escriben en el informe) y modo rapido ---
Public ModoRapido As Boolean           ' lo pone la ventana de revision (casilla "Modo rapido")
Private tBusca As Double               ' s buscando la cara de cada taladro
Private tCrea As Double                ' s creando taladros (croquis + corte + volumen)
Private tVol As Double                 ' s calculando volumenes (incluido en tCrea)
Private volPrevio As Double            ' volumen de la pieza tras el ultimo corte verificado (m3)
Private volValido As Boolean           ' True solo si volPrevio sigue siendo el volumen actual
Private rapidoActivo As Boolean
Private rapidoDoc As SldWorks.ModelDoc2

Private Function Dur(ByVal t0 As Double) As Double
    Dur = Timer - t0
    If Dur < 0 Then Dur = Dur + 86400      ' paso por medianoche
End Function

Private Function TextoTiempos(ByVal tIni As Double) As String
    TextoTiempos = "TIEMPOS (s): total " & Format(Dur(tIni), "0.0") & " | busqueda de cara " & Format(tBusca, "0.0") & _
                   " | creacion de taladros " & Format(tCrea, "0.0") & " (de ellos, calculo de volumen " & Format(tVol, "0.0") & ")" & _
                   " | modo rapido " & IIf(ModoRapido, "SI", "NO")
End Function

' Modo rapido: mientras se taladra una pieza no se repinta la pantalla ni se actualiza el arbol.
' SIEMPRE se restaura al terminar la pieza (RapidoOff) y en el manejador de errores.
Private Sub RapidoOn(md As SldWorks.ModelDoc2)
    If Not ModoRapido Then Exit Sub
    On Error Resume Next
    md.ActiveView.EnableGraphicsUpdate = False
    md.FeatureManager.EnableFeatureTree = False
    Set rapidoDoc = md
    rapidoActivo = True
End Sub

Private Sub RapidoOff()
    If Not rapidoActivo Then Exit Sub
    On Error Resume Next
    rapidoDoc.ActiveView.EnableGraphicsUpdate = True
    rapidoDoc.FeatureManager.EnableFeatureTree = True
    rapidoDoc.GraphicsRedraw2
    rapidoActivo = False
    Set rapidoDoc = Nothing
End Sub

Private Sub Anota(ByVal s As String)
    inf = inf & s & vbCrLf
End Sub

Public Sub Mecaniza()
    Dim app As SldWorks.SldWorks
    Set app = Application.SldWorks
    Dim amd As SldWorks.ModelDoc2
    Set amd = app.ActiveDoc
    If amd Is Nothing Then Exit Sub
    If amd.GetType = swDocPART Then MecanizaPieza app, amd: Exit Sub     ' pieza multicuerpo
    If amd.GetType <> swDocASSEMBLY Then MsgBox "Activa el ensamblaje.", vbExclamation, "Tarugos": Exit Sub
    ' El ensamblaje activo debe ser el que se analizo en la revision: la ventana de revision
    ' no es modal y, si se cambio de documento, Piezas()/Uniones() son de otro ensamblaje
    ' (se mezclarian rutas, copias de seguridad y sustituciones).
    If amd.GetPathName & "|" & amd.GetTitle <> mMain.AsmAnalizado Then
        MsgBox "El documento activo no es el ensamblaje que se analizo en la revision." & vbCrLf & _
               "Activa de nuevo ese ensamblaje y repite el analisis.", vbExclamation, "Tarugos"
        Exit Sub
    End If
    Dim asm As SldWorks.AssemblyDoc
    Set asm = amd
    inf = ""
    Dim carpeta As String, hechos As Long, nSustituidas As Long, nRepetidos As Long
    On Error GoTo errores

    CreaJobs
    If NJobs = 0 Then MsgBox "No hay tarugos que mecanizar.", vbInformation, "Tarugos": Exit Sub

    ' ---- piezas afectadas ----
    Dim nPz As Long, lista() As Long, i As Long, p As Long
    ReDim lista(0 To NJobs)
    For i = 0 To NJobs - 1
        If Not EnLista(lista, nPz, Jobs(i).Pieza) Then lista(nPz) = Jobs(i).Pieza: nPz = nPz + 1
    Next i

    ' modo: 0 omitir | 1 modificar el original | 2 instancia con copia independiente
    Dim modo() As Long, rutaDoc() As String, miemb() As String
    ReDim modo(0 To nPz): ReDim rutaDoc(0 To nPz): ReDim miemb(0 To nPz)
    Dim md As SldWorks.ModelDoc2, a As Long, b As Long
    For i = 0 To nPz - 1
        p = lista(i)
        modo(i) = 1
        Set md = Piezas(p).Comp.GetModelDoc2
        If md Is Nothing Then
            modo(i) = 0: Anota "OMITIDA " & Piezas(p).Nombre & ": sin documento"
        ElseIf Piezas(p).Comp.IsVirtual <> 0 Then
            modo(i) = 0
            Anota "OMITIDA " & Piezas(p).Nombre & ": pieza virtual (guardada dentro del ensamblaje). " & _
                  "Guardala como archivo externo (clic derecho > Guardar pieza (en archivo externo)) y repite."
        ElseIf md.GetPathName = "" Then
            modo(i) = 0: Anota "OMITIDA " & Piezas(p).Nombre & ": sin archivo propio"
        ElseIf mGuardado.EsSoloLectura(md) Then
            modo(i) = 0: Anota "OMITIDA " & Piezas(p).Nombre & ": archivo de solo lectura (sin extraer / protegido)"
        Else
            rutaDoc(i) = md.GetPathName
        End If
    Next i

    ' ---- piezas compartidas: varias instancias del mismo archivo ----
    Dim otras As String, todasIguales As Boolean, resp As VbMsgBoxResult
    Dim vv() As String, q As Long
    For i = 0 To nPz - 1
        If modo(i) = 1 Then
            otras = "": todasIguales = True
            For b = i + 1 To nPz - 1
                If modo(b) = 1 Then
                    If LCase(rutaDoc(b)) = LCase(rutaDoc(i)) Then
                        otras = otras & b & ","
                        If Not MismoConjunto(lista(i), lista(b)) Then todasIguales = False
                    End If
                End If
            Next b
            If otras <> "" Then
                vv = Split(Left(otras, Len(otras) - 1), ",")
                If todasIguales Then
                    ' mismos taladros en todas las instancias: basta modificar el archivo una vez
                    For q = 0 To UBound(vv)
                        modo(CLng(vv(q))) = 0
                        Anota "OK    " & Piezas(lista(CLng(vv(q)))).Nombre & ": instancia de " & Piezas(lista(i)).Nombre & _
                              " con los mismos taladros (se mecaniza el archivo una sola vez)"
                        For a = 0 To NJobs - 1
                            If Jobs(a).Pieza = lista(CLng(vv(q))) Then nRepetidos = nRepetidos + 1
                        Next a
                    Next q
                Else
                    resp = MsgBox("La pieza " & Mid(rutaDoc(i), InStrRev(rutaDoc(i), "\") + 1) & " esta repetida en el ensamblaje y " & _
                           "las instancias necesitan taladros DISTINTOS." & vbCrLf & vbCrLf & _
                           "SI = modificar el archivo original (afecta a todas las instancias: se crean los taladros de todas)" & vbCrLf & _
                           "NO = crear una copia independiente para cada instancia extra y sustituirla en este ensamblaje" & vbCrLf & _
                           "CANCELAR = no mecanizar esta pieza", vbYesNoCancel + vbQuestion, "Tarugos - pieza compartida")
                    If resp = vbYes Then
                        miemb(i) = Left(otras, Len(otras) - 1)
                        For q = 0 To UBound(vv): modo(CLng(vv(q))) = 0: Next q
                        Anota "AVISO " & Piezas(lista(i)).Nombre & ": se modifica el archivo original; afecta a todas sus instancias"
                    ElseIf resp = vbNo Then
                        For q = 0 To UBound(vv): modo(CLng(vv(q))) = 2: Next q
                    Else
                        modo(i) = 0
                        For q = 0 To UBound(vv): modo(CLng(vv(q))) = 0: Next q
                        Anota "OMITIDA " & Piezas(lista(i)).Nombre & ": cancelado por el usuario (pieza compartida)"
                    End If
                End If
            End If
        End If
    Next i

    ' ---- resumen y confirmacion ----
    Dim resumen As String, n As Long
    For i = 0 To nPz - 1
        If modo(i) > 0 Then
            n = 0
            For a = 0 To NJobs - 1
                If Jobs(a).Pieza = lista(i) Then n = n + 1
            Next a
            If modo(i) = 1 Then
                resumen = resumen & "  MODIFICAR  " & rutaDoc(i) & "  (" & n & " taladros)" & vbCrLf
            Else
                resumen = resumen & "  COPIA NUEVA de " & Mid(rutaDoc(i), InStrRev(rutaDoc(i), "\") + 1) & " para " & Piezas(lista(i)).Nombre & _
                          "  (" & n & " taladros)" & vbCrLf
            End If
        End If
    Next i
    If resumen = "" Then
        MsgBox "Ninguna pieza se puede mecanizar:" & vbCrLf & inf, vbExclamation, "Tarugos"
        Exit Sub
    End If
    Dim cerca As String
    cerca = TaladrosCercanos()
    If cerca <> "" Then Anota cerca
    If MsgBox("Se van a modificar / crear estos archivos:" & vbCrLf & vbCrLf & resumen & vbCrLf & _
              IIf(cerca <> "", cerca & vbCrLf, "") & _
              "Diametro " & Cfg.Diametro & " mm. Antes se hara una copia de seguridad de los originales." & vbCrLf & _
              "Los taladros que ya existan no se duplican. Los tarugos DESCARTADOS que ya estuvieran taladrados se ELIMINAN de las piezas." & vbCrLf & "Â¿Continuar?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Tarugos - mecanizado") <> vbYes Then
        Exit Sub
    End If

    ' ---- copia de seguridad de los originales que se van a modificar ----
    Dim destino As String
    carpeta = mGuardado.CarpetaRespaldo(amd)
    If carpeta = "" Then
        MsgBox "No se pudo crear la carpeta de copia de seguridad junto al ensamblaje (permisos o disco)." & vbCrLf & _
               "No se ha modificado nada.", vbExclamation, "Tarugos"
        Exit Sub
    End If
    For i = 0 To nPz - 1
        If modo(i) = 1 Then
            Set md = Piezas(lista(i)).Comp.GetModelDoc2
            destino = carpeta & "\" & Mid(rutaDoc(i), InStrRev(rutaDoc(i), "\") + 1)
            If Not mGuardado.CopiaA(md, destino) Then
                modo(i) = 0
                Anota "OMITIDA " & Piezas(lista(i)).Nombre & ": no se pudo hacer la copia de seguridad"
            End If
        End If
    Next i
    Anota "Copia de seguridad en: " & carpeta
    ' copia tambien el archivo del propio ensamblaje (las piezas se respaldan arriba)
    If Not mGuardado.CopiaA(amd, carpeta & "\" & Mid(amd.GetPathName, InStrRev(amd.GetPathName, "\") + 1)) Then
        Anota "AVISO no se pudo copiar el ensamblaje a la carpeta de respaldo"
    End If

    ' ---- copias independientes (se hacen ANTES de modificar ningun original) ----
    Dim nueva As String
    For i = 0 To nPz - 1
        If modo(i) = 2 Then
            Set md = Piezas(lista(i)).Comp.GetModelDoc2
            nueva = mGuardado.RutaCopiaLibre(rutaDoc(i))
            If mGuardado.CopiaA(md, nueva) Then
                rutaDoc(i) = nueva
                Anota "Copia independiente creada: " & nueva
            Else
                modo(i) = 0
                Anota "ERROR " & Piezas(lista(i)).Nombre & ": no se pudo crear la copia independiente"
            End If
        End If
    Next i

    ' ---- mecanizar documento a documento ----
    Dim titAsm As String
    titAsm = amd.GetTitle
    nSaltados = 0: nBorrados = 0
    Dim tIni As Double
    tIni = Timer: tBusca = 0: tCrea = 0: tVol = 0
    Dim miemb2() As Long, nmm As Long, vm() As String
    For i = 0 To nPz - 1
        If modo(i) > 0 Then
            ReDim miemb2(0 To nPz + 1)
            miemb2(0) = lista(i): nmm = 1
            If miemb(i) <> "" Then
                vm = Split(miemb(i), ",")
                For q = 0 To UBound(vm)
                    miemb2(nmm) = lista(CLng(vm(q))): nmm = nmm + 1
                Next q
            End If
            hechos = hechos + MecanizaDoc(app, rutaDoc(i), miemb2, nmm)
        End If
    Next i

    ' ---- volver al ensamblaje y sustituir las instancias por sus copias ----
    Dim er As Long
    app.ActivateDoc3 titAsm, False, 1, er
    Set amd = app.ActiveDoc
    Set asm = amd
    For i = 0 To nPz - 1
        If modo(i) = 2 Then
            If mGuardado.ReemplazaInstancia(asm, amd, Piezas(lista(i)).Comp, rutaDoc(i)) Then
                Anota "Instancia " & Piezas(lista(i)).Nombre & " sustituida por " & rutaDoc(i)
                nSustituidas = nSustituidas + 1
            Else
                Anota "ERROR no se pudo sustituir " & Piezas(lista(i)).Nombre & " por su copia; hazlo manualmente (Reemplazar componentes)"
            End If
        End If
    Next i
    asm.ForceRebuild3 False

    Dim ruta As String
    Anota TextoTiempos(tIni)
    ruta = EscribeInformeFinal(carpeta, amd, hechos)
    MsgBox "Taladros creados: " & hechos & " de " & NJobs & vbCrLf & "Ya existentes (no repetidos): " & nSaltados & vbCrLf & "Eliminados (descartados): " & nBorrados & vbCrLf & _
           IIf(nRepetidos > 0, vbCrLf & "Aclaracion: " & nRepetidos & " de los " & NJobs & " taladros son de piezas REPETIDAS (mismo archivo y mismos taladros)." & vbCrLf & _
               "Ya estan hechos en el archivo y por eso no se cuentan aparte." & vbCrLf, "") & _
           IIf(nSustituidas > 0, vbCrLf & "El ensamblaje ha cambiado (" & nSustituidas & " instancia(s) sustituida(s)): GUARDALO, no se guarda solo." & vbCrLf, "") & _
           vbCrLf & "Tiempo de mecanizado: " & Format(Dur(tIni), "0") & " s" & vbCrLf & "Informe: " & ruta, vbInformation, "Tarugos - Fase 5"
    Exit Sub

errores:
    ' Error no previsto: no se pierde el informe. Las piezas ya guardadas lo estan; la que
    ' se estaba procesando puede quedar abierta SIN guardar (hay copia de seguridad).
    Dim txtErr As String
    txtErr = Err.Number & " - " & Err.Description        ' se guarda antes: On Error limpia Err
    RapidoOff
    Anota "ERROR INESPERADO: " & txtErr & _
          ". Las piezas ya guardadas lo estan; la que se procesaba puede haber quedado abierta sin guardar. Revisa el informe y la copia de seguridad."
    On Error Resume Next
    ruta = EscribeInformeFinal(carpeta, amd, hechos)
    MsgBox "Error inesperado durante el mecanizado: " & txtErr & vbCrLf & vbCrLf & _
           "Informe: " & ruta, vbCritical, "Tarugos - Fase 5"
End Sub

' Tareas: un taladro por tarugo y pieza. Activo = crear (si no existe);
' descartado = ELIMINAR el taladro si ya existia (lo hizo esta macro)
Private Sub CreaJobs()
    NJobs = 0
    Dim u As Long, k As Long, j As TJob
    For u = 0 To NUniones - 1
        For k = 0 To Uniones(u).NT - 1
            j.U = u: j.K = k: j.Pc = Uniones(u).Pos(k)
            j.Borrar = Not Uniones(u).Activo(k)
            ' pieza 1: cara con normal N1, profundidad Prof1
            j.Pieza = Uniones(u).P1: j.Nout = Uniones(u).N1: j.Prof = Cfg.Prof1 * MM
            AnadeJob j
            ' pieza 2: cara con normal -N1, profundidad Prof2
            j.Pieza = Uniones(u).P2
            j.Nout = V3(-Uniones(u).N1.X, -Uniones(u).N1.Y, -Uniones(u).N1.Z)
            j.Prof = Cfg.Prof2 * MM
            AnadeJob j
        Next k
    Next u
End Sub

' ---------------------------------------------------------------------------
' PIEZA MULTICUERPO: cada cuerpo solido es una pieza y todos los taladros van en el MISMO
' archivo (cada corte se limita a su cuerpo). Mismo flujo que el ensamblaje, sin piezas
' repetidas ni copias independientes.
' ---------------------------------------------------------------------------
Private Sub MecanizaPieza(app As SldWorks.SldWorks, amd As SldWorks.ModelDoc2)
    If amd.GetPathName & "|" & amd.GetTitle <> mMain.AsmAnalizado Then
        MsgBox "El documento activo no es la pieza que se analizo en la revision." & vbCrLf & _
               "Activala de nuevo y repite el analisis.", vbExclamation, "Tarugos"
        Exit Sub
    End If
    If amd.GetPathName = "" Then MsgBox "Guarda la pieza antes de mecanizar.", vbExclamation, "Tarugos": Exit Sub
    If mGuardado.EsSoloLectura(amd) Then MsgBox "La pieza es de solo lectura (sin extraer / protegida).", vbExclamation, "Tarugos": Exit Sub
    inf = ""
    Dim carpeta As String, hechos As Long
    Dim tIni As Double
    tIni = Timer
    On Error GoTo errores

    CreaJobs
    If NJobs = 0 Then MsgBox "No hay tarugos que mecanizar.", vbInformation, "Tarugos": Exit Sub

    ' piezas (cuerpos) afectadas = miembros del unico documento
    Dim nPz As Long, lista() As Long, i As Long
    ReDim lista(0 To NJobs)
    For i = 0 To NJobs - 1
        If Not EnLista(lista, nPz, Jobs(i).Pieza) Then lista(nPz) = Jobs(i).Pieza: nPz = nPz + 1
    Next i

    Dim cerca As String
    cerca = TaladrosCercanos()
    If cerca <> "" Then Anota cerca
    If MsgBox("Se va a MODIFICAR esta pieza (todos los taladros en el mismo archivo):" & vbCrLf & vbCrLf & _
              "  " & amd.GetPathName & "  (" & NJobs & " taladros en " & nPz & " cuerpos)" & vbCrLf & vbCrLf & _
              IIf(cerca <> "", cerca & vbCrLf, "") & _
              "Diametro " & Cfg.Diametro & " mm. Antes se hara una copia de seguridad de la pieza." & vbCrLf & _
              "Los taladros que ya existan no se duplican. Los tarugos DESCARTADOS que ya estuvieran taladrados se ELIMINAN." & vbCrLf & "¿Continuar?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Tarugos - mecanizado") <> vbYes Then
        Exit Sub
    End If

    carpeta = mGuardado.CarpetaRespaldo(amd)
    If carpeta = "" Then
        MsgBox "No se pudo crear la carpeta de copia de seguridad junto a la pieza (permisos o disco)." & vbCrLf & _
               "No se ha modificado nada.", vbExclamation, "Tarugos"
        Exit Sub
    End If
    If Not mGuardado.CopiaA(amd, carpeta & "\" & Mid(amd.GetPathName, InStrRev(amd.GetPathName, "\") + 1)) Then
        Anota "ERROR no se pudo hacer la copia de seguridad: no se mecaniza"
    Else
        Anota "Copia de seguridad en: " & carpeta
        nSaltados = 0: nBorrados = 0
        tIni = Timer: tBusca = 0: tCrea = 0: tVol = 0
        hechos = MecanizaDoc(app, amd.GetPathName, lista, nPz)
    End If

    Dim ruta As String
    Anota TextoTiempos(tIni)
    ruta = EscribeInformeFinal(carpeta, amd, hechos)
    MsgBox "Taladros creados: " & hechos & " de " & NJobs & vbCrLf & "Ya existentes (no repetidos): " & nSaltados & vbCrLf & "Eliminados (descartados): " & nBorrados & vbCrLf & _
           vbCrLf & "Tiempo de mecanizado: " & Format(Dur(tIni), "0") & " s" & vbCrLf & "Informe: " & ruta, vbInformation, "Tarugos - Fase 5"
    Exit Sub

errores:
    Dim txtErr As String
    txtErr = Err.Number & " - " & Err.Description
    RapidoOff
    Anota "ERROR INESPERADO: " & txtErr
    On Error Resume Next
    ruta = EscribeInformeFinal(carpeta, amd, hechos)
    MsgBox "Error inesperado durante el mecanizado: " & txtErr & vbCrLf & vbCrLf & "Informe: " & ruta, vbCritical, "Tarugos - Fase 5"
End Sub

' Escribe Informe_Fase5.txt y el CSV. Si el error ocurrio antes de crear la carpeta de
' respaldo, se usa la carpeta del ensamblaje. Devuelve la ruta del informe.
Private Function EscribeInformeFinal(ByVal carpeta As String, amd As SldWorks.ModelDoc2, ByVal hechos As Long) As String
    If carpeta = "" Then carpeta = Left(amd.GetPathName, InStrRev(amd.GetPathName, "\") - 1)
    Dim ruta As String
    ruta = carpeta & "\Informe_Fase5.txt"
    Dim ff As Integer
    ff = FreeFile
    Open ruta For Output As #ff
    Print #ff, "TARUGOS PRO - Fases 4/5 - Taladros"
    Print #ff, "Ensamblaje: " & amd.GetPathName & "   " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #ff, "Taladros creados: " & hechos & " de " & NJobs & "   (ya existentes, no repetidos: " & nSaltados & ")   Eliminados: " & nBorrados
    Print #ff, String(60, "-")
    Print #ff, inf
    Close #ff
    mInformes.ExportaCSV carpeta & "\Informe_Tarugos.csv", amd.GetTitle
    EscribeInformeFinal = ruta
End Function

' Taladros de la pieza p en coordenadas de la PIEZA
Private Function JobsPieza(ByVal p As Long, pps() As TV3, nns() As TV3, profs() As Double) As Long
    Dim inv As SldWorks.MathTransform, a As Long, n As Long
    Set inv = Piezas(p).Comp.GetTotalTransform(False).Inverse
    ReDim pps(0 To NJobs): ReDim nns(0 To NJobs): ReDim profs(0 To NJobs)
    For a = 0 To NJobs - 1
        If Jobs(a).Pieza = p Then
            pps(n) = TrP(inv, Jobs(a).Pc)
            nns(n) = VUnit(TrV(inv, Jobs(a).Nout))
            profs(n) = Jobs(a).Prof
            n = n + 1
        End If
    Next a
    JobsPieza = n
End Function

' True si las instancias pa y pb necesitan exactamente los mismos taladros (en coord. de la pieza)
Private Function MismoConjunto(ByVal pa As Long, ByVal pb As Long) As Boolean
    Dim p1() As TV3, n1() As TV3, d1() As Double, p2() As TV3, n2() As TV3, d2() As Double
    Dim na As Long, nb As Long, i As Long, j As Long, hallado As Boolean
    na = JobsPieza(pa, p1, n1, d1)
    nb = JobsPieza(pb, p2, n2, d2)
    If na <> nb Then Exit Function
    For i = 0 To na - 1
        hallado = False
        For j = 0 To nb - 1
            If VLen(VSub(p1(i), p2(j))) < 0.0002 And VDot(n1(i), n2(j)) > 0.99 And Abs(d1(i) - d2(j)) < 0.00001 Then hallado = True: Exit For
        Next j
        If Not hallado Then Exit Function
    Next i
    MismoConjunto = True
End Function

'------------------------------------------------------------------------------
' Abre el documento, crea los taladros que falten (de las instancias "miembros"),
' guarda UNA vez. Devuelve cuantos taladros se crearon.
'------------------------------------------------------------------------------
Private Function MecanizaDoc(app As SldWorks.SldWorks, ByVal path As String, miembros() As Long, ByVal nm As Long) As Long
    Dim e As Long, w As Long
    Dim pmd As SldWorks.ModelDoc2
    Dim creados As Long                 ' (se declara aqui para que el manejador de errores lo vea)
    Dim t1 As Double
    volValido = False                   ' el volumen guardado es de otra pieza
    On Error GoTo errDoc
    ' Si la pieza ya tenia ventana abierta (la abrio el usuario), no se cierra al terminar:
    ' solo se cierran las ventanas que abre la macro. Una pieza cargada por el ensamblaje
    ' pero sin ventana tiene Visible = False.
    Dim ventanaPrevia As Boolean
    Set pmd = app.GetOpenDocumentByName(path)
    If Not pmd Is Nothing Then ventanaPrevia = pmd.Visible
    Set pmd = Nothing
    ' El documento ACTIVO nunca se cierra (pieza multicuerpo: es el propio documento analizado;
    ' si se cerrara, el informe final y el resto de la macro trabajarian con un documento invalido)
    Set pmd = app.ActiveDoc
    If Not pmd Is Nothing Then If LCase(pmd.GetPathName) = LCase(path) Then ventanaPrevia = True
    Set pmd = Nothing
    ' La pieza se abre Y SE ACTIVA (con ventana). El modo sin ventana (DocumentVisible False,
    ' sin ActivateDoc3) fallaba: SelectByID2 no devolvia la cara y InsertSketch no abria el
    ' croquis (0 de 118 taladros), porque SolidWorks necesita el documento activo para
    ' seleccionar y croquizar.
    Set pmd = app.OpenDoc6(path, swDocPART, swOpenDocOptions_Silent, "", e, w)
    If pmd Is Nothing Then Anota "ERROR no se pudo abrir " & path: Exit Function
    Dim eAct As Long
    app.ActivateDoc3 pmd.GetTitle, False, 1, eAct
    Set pmd = app.ActiveDoc
    If pmd Is Nothing Then Anota "ERROR no se pudo activar " & path: Exit Function
    Dim act As SldWorks.ModelDoc2
    Set act = pmd
    If mGuardado.EsSoloLectura(act) Then Anota "OMITIDA " & path & ": solo lectura": Exit Function

    Dim a As Long, m As Long, pp As TV3, nn As TV3, f As SldWorks.Face2
    Dim inv As SldWorks.MathTransform
    Dim ex() As mDuplicados.TMeta, nex As Long, est As Long, profmm As Double
    Dim nombreEx As String, borrados As Long
    hayAviso = False
    RapidoOn act
    Anota "   existentes en " & act.GetTitle & ":" & vbCrLf & mDuplicados.Diag(act)
    For m = 0 To nm - 1
        ' pieza multicuerpo (Comp = Nothing): las coordenadas del documento ya son las de la pieza
        If Piezas(miembros(m)).Comp Is Nothing Then Set inv = Nothing Else Set inv = Piezas(miembros(m)).Comp.GetTotalTransform(False).Inverse
        For a = 0 To NJobs - 1
            If Jobs(a).Pieza = miembros(m) Then
                If inv Is Nothing Then
                    pp = Jobs(a).Pc
                    nn = VUnit(Jobs(a).Nout)
                Else
                    pp = TrP(inv, Jobs(a).Pc)
                    nn = VUnit(TrV(inv, Jobs(a).Nout))
                End If
                profmm = Jobs(a).Prof / MM
                ex = mDuplicados.LeeExistentes(act, nex)
                If Jobs(a).Borrar Then
                    nombreEx = mDuplicados.NombreExistente(ex, nex, pp, nn)
                    volValido = False           ' borrar una operacion cambia el volumen
                    If nombreEx <> "" Then
                        If BorraOperacion(act, nombreEx) Then
                            nBorrados = nBorrados + 1: borrados = borrados + 1
                            Anota "ELIMINADO " & nombreEx & " (tarugo descartado)"
                        Else
                            Anota "ERROR no se pudo eliminar " & nombreEx
                        End If
                    End If
                    GoTo siguiente
                End If
                est = mDuplicados.EstadoTaladro(ex, nex, pp, nn, Cfg.Diametro, profmm)
                If est = 1 Then
                    nSaltados = nSaltados + 1
                    Anota "YA EXISTE " & Jobs(a).Nombre & " (no se duplica)"
                ElseIf est = 2 Then
                    nSaltados = nSaltados + 1
                    Anota "AVISO " & Jobs(a).Nombre & ": ya hay un taladro de la macro en ese punto con otro diametro/profundidad; no se modifica"
                Else
                    t1 = Timer
                    Set f = BuscaCara(act, pp, nn, Jobs(a).Pieza)
                    tBusca = tBusca + Dur(t1)
                    If f Is Nothing Then
                        Anota "ERROR " & Jobs(a).Nombre & ": no se encontro la cara de contacto en la pieza"
                    ElseIf mUniones.AgujeroEnCara(f, pp) Then
                        ' el modelo ya trae ahi un agujero del diametro del tarugo: no se taladra otro encima
                        nSaltados = nSaltados + 1
                        Anota "YA HAY un agujero del modelo en " & Jobs(a).Nombre & " (no se taladra otro encima)"
                    Else
                        t1 = Timer
                        If CreaTaladro(act, f, pp, nn, Jobs(a)) Then creados = creados + 1
                        tCrea = tCrea + Dur(t1)
                    End If
                End If
            End If
siguiente:
        Next a
    Next m
    RapidoOff                               ' la pantalla y el arbol vuelven a la normalidad ANTES de agrupar y guardar

    If hayAviso Then
        Anota "NO GUARDADA " & path & ": algun taladro elimino un volumen distinto del esperado. " & _
              "La pieza queda modificada solo en memoria: cierrala SIN guardar (hay copia de seguridad)."
        Exit Function
    End If
    If creados > 0 Then AgrupaEnCarpeta act
    If creados + borrados > 0 Then
        Dim er As Long
        er = mGuardado.GuardaDoc(act)
        If er <> 0 Then
            ' Si no se guardo, la ventana se deja abierta para no perder los cambios
            Anota "AVISO error al guardar " & path & " (codigo " & er & ")"
        Else
            Anota "Guardada " & path & " (" & creados & " creados, " & borrados & " eliminados)"
            ' Ya guardada: se cierra su ventana para que no queden una por pieza abiertas
            ' (y no pregunte "guardar cambios" al cerrar). Si el ensamblaje la usa, el
            ' documento sigue cargado en memoria; solo se cierra la ventana.
            If Not ventanaPrevia Then app.CloseDoc act.GetTitle
        End If
    Else
        Anota "Sin cambios en " & path
        If Not ventanaPrevia Then app.CloseDoc act.GetTitle
    End If
    MecanizaDoc = creados
    Exit Function

errDoc:
    RapidoOff
    ' Error no previsto en esta pieza: se anota y se sigue con las demas. La ventana queda
    ' abierta SIN guardar (no se sabe en que estado esta); hay copia de seguridad.
    Anota "ERROR INESPERADO en " & path & ": " & Err.Number & " - " & Err.Description & _
          ". La pieza queda abierta sin guardar: revisala (hay copia de seguridad)."
    MecanizaDoc = creados
End Function

' Pares de taladros ACTIVOS de la misma pieza, de uniones distintas, cuyos ejes se cruzan o
' pasan a menos de 1,5 diametros (por ejemplo, tarugos de dos uniones perpendiculares que
' coinciden en una esquina). Solo es un aviso: no cambia nada. Devuelve "" si no hay ninguno.
Private Function TaladrosCercanos() As String
    Dim a As Long, b As Long, n As Long, r As String
    For a = 0 To NJobs - 2
        If Not Jobs(a).Borrar Then
            For b = a + 1 To NJobs - 1
                If Jobs(b).Pieza = Jobs(a).Pieza Then
                    If Not Jobs(b).Borrar And Jobs(b).U <> Jobs(a).U Then
                        If DistanciaEjes(a, b) < 1.5 * Cfg.Diametro * MM Then
                            n = n + 1
                            If n <= 10 Then r = r & "   " & Jobs(a).Nombre & "  y  " & Jobs(b).Nombre & vbCrLf
                        End If
                    End If
                End If
            Next b
        End If
    Next a
    If n > 0 Then
        TaladrosCercanos = "AVISO: " & n & " par(es) de taladros de la misma pieza se cruzan o quedan a menos de 1,5 diametros:" & _
                           vbCrLf & r & IIf(n > 10, "   ..." & vbCrLf, "")
    End If
End Function

' Distancia minima entre los ejes de dos taladros, midiendo 8 puntos en cada eje
' (del punto de contacto hacia dentro de la pieza, una longitud igual a la profundidad)
Private Function DistanciaEjes(ByVal a As Long, ByVal b As Long) As Double
    Dim i As Long, j As Long, t As Double, u As Double, d As Double, mejor As Double
    Dim pa As TV3, pb As TV3
    mejor = 1E+30
    For i = 0 To 7
        t = i / 7 * Jobs(a).Prof
        pa = V3(Jobs(a).Pc.X - Jobs(a).Nout.X * t, Jobs(a).Pc.Y - Jobs(a).Nout.Y * t, Jobs(a).Pc.Z - Jobs(a).Nout.Z * t)
        For j = 0 To 7
            u = j / 7 * Jobs(b).Prof
            pb = V3(Jobs(b).Pc.X - Jobs(b).Nout.X * u, Jobs(b).Pc.Y - Jobs(b).Nout.Y * u, Jobs(b).Pc.Z - Jobs(b).Nout.Z * u)
            d = VLen(VSub(pa, pb))
            If d < mejor Then mejor = d
        Next j
    Next i
    DistanciaEjes = mejor
End Function

Private Sub AnadeJob(j As TJob)
    j.Nombre = Uniones(j.U).Id & "_TARUGO_" & Format(j.K + 1, "000") & "_" & Limpio(Piezas(j.Pieza).Nombre) & _
               "_D" & Cfg.Diametro & "_P" & Cfg.Prof1 & "-" & Cfg.Prof2
    ReDim Preserve Jobs(0 To NJobs)
    Jobs(NJobs) = j
    NJobs = NJobs + 1
End Sub

' Elimina una operacion de la macro (con su croquis absorbido)
Private Function BorraOperacion(md As SldWorks.ModelDoc2, ByVal nombre As String) As Boolean
    Dim pd As SldWorks.PartDoc, ft As SldWorks.Feature
    Set pd = md
    Set ft = pd.FeatureByName(nombre)
    If ft Is Nothing Then Exit Function
    md.ClearSelection2 True
    ft.Select2 False, 0
    BorraOperacion = md.Extension.DeleteSelection2(swDelete_Absorbed)
    md.ClearSelection2 True
End Function

Private Function Limpio(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c Like "[A-Za-z0-9_-]" Then r = r & c Else r = r & "_"
    Next i
    Limpio = r
End Function

Private Function EnLista(l() As Long, ByVal n As Long, ByVal v As Long) As Boolean
    Dim i As Long
    For i = 0 To n - 1
        If l(i) = v Then EnLista = True: Exit Function
    Next i
End Function

' Cara plana de la pieza (espacio de la pieza) con normal nn que contiene el punto pp
' Primero se busca la cara con tolerancia estricta; si no aparece, otra pasada mas holgada
' (la cara puede quedar partida o con la caja algo menor que el punto, p. ej. junto a un agujero).
Private Function BuscaCara(pmd As SldWorks.ModelDoc2, pp As TV3, nn As TV3, ByVal pieza As Long) As SldWorks.Face2
    Set BuscaCara = BuscaCaraTol(pmd, pp, nn, pieza, 0.0002)
    If BuscaCara Is Nothing Then Set BuscaCara = BuscaCaraTol(pmd, pp, nn, pieza, 0.003)
End Function

Private Function BuscaCaraTol(pmd As SldWorks.ModelDoc2, pp As TV3, nn As TV3, ByVal pieza As Long, ByVal tolBox As Double) As SldWorks.Face2
    Dim bodies As Variant
    bodies = CType2(pmd).GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) Then Exit Function
    Dim b As Long, f As SldWorks.Face2, bd As SldWorks.Body2
    Dim s As SldWorks.Surface, nf As TV3, pf As TV3, v As Variant, bx As Variant
    Dim tolP As Double
    tolP = MaxD(tolBox, Cfg.Tol * MM + 0.00001)      ' las piezas pueden estar separadas hasta la holgura configurada
    For b = 0 To UBound(bodies)
        Set bd = bodies(b)
        ' pieza multicuerpo: solo las caras del cuerpo que ES la pieza (dos cuerpos pueden tener
        ' caras coincidentes en el mismo plano)
        If Not CuerpoEsPieza(bd, pieza) Then GoTo siguienteCuerpo
        Set f = bd.GetFirstFace
        Do While Not f Is Nothing
            Set s = f.GetSurface
            If s.Identity = 4001 Then
                v = f.Normal
                nf = V3(v(0), v(1), v(2))
                If VDot(nf, nn) > 0.999 Then
                    v = s.PlaneParams
                    pf = V3(v(3), v(4), v(5))
                    If Abs(VDot(VSub(pf, pp), nn)) < tolP Then
                        bx = f.GetBox
                        If pp.X >= bx(0) - tolBox And pp.X <= bx(3) + tolBox And _
                           pp.Y >= bx(1) - tolBox And pp.Y <= bx(4) + tolBox And _
                           pp.Z >= bx(2) - tolBox And pp.Z <= bx(5) + tolBox Then
                            Set BuscaCaraTol = f
                            Exit Function
                        End If
                    End If
                End If
            End If
            Set f = f.GetNextFace
        Loop
siguienteCuerpo:
    Next b
End Function

' True si el cuerpo bd es la pieza p. En ensamblajes (Cuerpo = Nothing) vale cualquier cuerpo, como antes.
' Se compara la CAJA del cuerpo (SolidWorks renombra el cuerpo al cortarlo, el nombre no sirve).
Private Function CuerpoEsPieza(ByVal bd As SldWorks.Body2, ByVal p As Long) As Boolean
    If Piezas(p).Cuerpo Is Nothing Then CuerpoEsPieza = True: Exit Function
    Dim bx As Variant
    bx = bd.GetBodyBox
    If IsEmpty(bx) Then Exit Function
    CuerpoEsPieza = (Abs(bx(0) - Piezas(p).BMin.X) < 0.0005 And Abs(bx(1) - Piezas(p).BMin.Y) < 0.0005 And _
                     Abs(bx(2) - Piezas(p).BMin.Z) < 0.0005 And Abs(bx(3) - Piezas(p).BMax.X) < 0.0005 And _
                     Abs(bx(4) - Piezas(p).BMax.Y) < 0.0005 And Abs(bx(5) - Piezas(p).BMax.Z) < 0.0005)
End Function

' El cuerpo real (en el documento md, tal como esta ahora) que corresponde a la pieza p, o Nothing
Private Function CuerpoDePieza(md As SldWorks.ModelDoc2, ByVal p As Long) As SldWorks.Body2
    If Piezas(p).Cuerpo Is Nothing Then Exit Function
    Dim bodies As Variant, b As Long
    bodies = CType2(md).GetBodies2(swSolidBody, False)
    If IsEmpty(bodies) Then Exit Function
    For b = 0 To UBound(bodies)
        If CuerpoEsPieza(bodies(b), p) Then Set CuerpoDePieza = bodies(b): Exit Function
    Next b
End Function

Private Function CType2(md As SldWorks.ModelDoc2) As SldWorks.PartDoc
    Set CType2 = md
End Function

Private Function Volumen(md As SldWorks.ModelDoc2) As Double
    Dim mp As SldWorks.MassProperty, t0 As Double
    t0 = Timer
    Set mp = md.Extension.CreateMassProperty
    If Not mp Is Nothing Then Volumen = mp.Volume
    tVol = tVol + Dur(t0)
End Function

'------------------------------------------------------------------------------
' Sketch circular en la cara + Corte-extruido ciego. Verifica el volumen eliminado.
'------------------------------------------------------------------------------
Private Function CreaTaladro(md As SldWorks.ModelDoc2, f As SldWorks.Face2, pp As TV3, nn As TV3, jb As TJob) As Boolean
    ' Volumen antes del corte: si el corte anterior de esta pieza se verifico, su volumen final ES el
    ' de ahora (nada mas ha cambiado): se ahorra un calculo de masa por taladro. Se invalida en cuanto
    ' algo puede haberlo cambiado (fallo, borrado de una operacion, otra pieza).
    Dim v0 As Double
    If volValido Then v0 = volPrevio Else v0 = Volumen(md)
    volValido = False

    ' El punto del tarugo esta en el plano de contacto de la pieza 1; la cara de ESTA pieza puede
    ' estar a unas decimas (holgura entre piezas). Se proyecta el punto sobre el plano de la cara.
    Dim pp0 As TV3, sfc As SldWorks.Surface, pv As Variant
    pp0 = pp
    Set sfc = f.GetSurface
    If Not sfc Is Nothing Then
        If sfc.Identity = 4001 Then
            pv = sfc.PlaneParams
            Dim dd As Double
            dd = VDot(VSub(V3(pv(3), pv(4), pv(5)), pp), nn)
            pp = V3(pp.X + nn.X * dd, pp.Y + nn.Y * dd, pp.Z + nn.Z * dd)
        End If
    End If

    ' Se selecciona la cara por su posicion (SelectByID2); Face2.Select4 con caras de
    ' GetBodies2 selecciono una cara equivocada en las pruebas. Se verifica la normal.
    md.ClearSelection2 True
    md.Extension.SelectByID2 "", "FACE", pp.X, pp.Y, pp.Z, False, 0, Nothing, 0
    Dim okSel As Boolean, sf As SldWorks.Face2, vn As Variant
    If md.SelectionManager.GetSelectedObjectCount2(-1) >= 1 Then
        If md.SelectionManager.GetSelectedObjectType3(1, -1) = swSelFACES Then
            Set sf = md.SelectionManager.GetSelectedObject6(1, -1)
            vn = sf.Normal
            okSel = (vn(0) * nn.X + vn(1) * nn.Y + vn(2) * nn.Z > 0.99)
        End If
    End If
    If Not okSel Then
        Anota "AVISO " & jb.Nombre & ": SelectByID2 no devolvio la cara esperada; se usa la cara encontrada en la pieza"
        md.ClearSelection2 True
        f.Select4 False, Nothing
    End If
    SnapshotCroquis md                 ' nombres de los croquis que ya hay (para reconocer el nuevo)
    md.SketchManager.InsertSketch True
    Dim sk As SldWorks.Sketch
    Set sk = md.SketchManager.ActiveSketch
    If sk Is Nothing Then
        Anota "ERROR " & jb.Nombre & ": no se pudo abrir el croquis": Exit Function
    End If
    Dim ps As TV3, m2s As SldWorks.MathTransform
    Set m2s = sk.ModelToSketchTransform
    ps = TrP(m2s, pp)
    ' un punto del plano del croquis debe quedar con z = 0 en coordenadas de croquis;
    ' si no, la transformacion devuelta va en el otro sentido (se elige la de menor |z|)
    If Abs(ps.Z) > 0.00001 Then
        Dim ps2 As TV3
        ps2 = TrP(m2s.Inverse, pp)
        If Abs(ps2.Z) < Abs(ps.Z) Then ps = ps2
    End If
    ' AddToDB: crear la entidad tal cual, SIN inferencias ni relaciones automaticas.
    ' Sin esto SolidWorks ajustaba el circulo a tangente con el borde de la cara
    ' (radio 9,5 mm en vez de 4 mm) y el corte fallaba por "geometria de espesor cero".
    md.SketchManager.AddToDB = True
    md.SketchManager.CreateCircleByRadius ps.X, ps.Y, 0, Cfg.Diametro * MM / 2
    md.SketchManager.AddToDB = False
    md.SketchManager.InsertSketch True

    Dim feat As SldWorks.Feature
    ' OJO: el parametro "Flip" de FeatureCut4 significa "cortar por FUERA del perfil".
    ' Para invertir el sentido del corte se usa "Dir".
    Dim cuerpoPieza As SldWorks.Body2
    Set cuerpoPieza = CuerpoDePieza(md, jb.Pieza)          ' Nothing salvo en piezas multicuerpo
    Set feat = CortaCiego(md, jb.Prof, False, cuerpoPieza)
    If feat Is Nothing Then Set feat = CortaCiego(md, jb.Prof, True, cuerpoPieza)      ' sentido contrario
    If feat Is Nothing Then
        Anota "ERROR " & jb.Nombre & ": el corte-extruido fallo (circulo en croquis: " & Format(ps.X * 1000, "0.0") & " ; " & _
              Format(ps.Y * 1000, "0.0") & " ; " & Format(ps.Z * 1000, "0.0") & " mm; croquis=" & NombreCroquisNuevo(md) & _
              "; cuerpo=" & IIf(cuerpoPieza Is Nothing, "no", "si") & ")"
        BorraUltimoCroquis md          ' no dejar croquis huerfanos en la pieza
        Exit Function
    End If

    Dim v1 As Double, esperado As Double, quitado As Double
    v1 = Volumen(md)
    volPrevio = v1: volValido = True
    quitado = v0 - v1
    esperado = 3.14159265358979 * (Cfg.Diametro * MM / 2) ^ 2 * jb.Prof
    feat.Name = jb.Nombre
    mDuplicados.Marca feat, pp0, nn, Cfg.Diametro, jb.Prof / MM
    Anota "      marca: " & mDuplicados.UltimoAdd & "  tipo=" & feat.GetTypeName2
    If Abs(quitado - esperado) > 0.05 * esperado Then
        ' Menos volumen del esperado (el agujero roza un hueco o el borde) es aceptable; mucho mas, o casi
        ' nada, indica que el corte ha ido a otro sitio y se bloquea el guardado.
        If quitado > 1.05 * esperado Or quitado < 0.3 * esperado Then hayAviso = True
        Anota "AVISO " & jb.Nombre & ": volumen eliminado " & Format(quitado * 1E+09, "0.0") & " mm3, esperado " & Format(esperado * 1E+09, "0.0") & " mm3"
    Else
        Anota "OK    " & jb.Nombre & ": " & Format(quitado * 1E+09, "0.0") & " mm3 (esperado " & Format(esperado * 1E+09, "0.0") & ")"
    End If
    CreaTaladro = True
End Function

' Mete los cortes de tarugo (UNION_xxx_TARUGO_xxx...) que estan sueltos en el arbol dentro de una
' carpeta "TARUGOS MACRO". Los croquis ya cuelgan de cada corte, asi que quedan dentro. Si falla solo se
' avisa: los taladros ya estan hechos y no se pierde nada.
Private Sub AgrupaEnCarpeta(md As SldWorks.ModelDoc2)
    On Error GoTo fallo
    Dim ft As SldWorks.Feature, n As Long, ok As Boolean
    md.ClearSelection2 True
    Set ft = md.FirstFeature
    Do While Not ft Is Nothing
        If Left(ft.Name, 6) = "UNION_" And InStr(ft.Name, "_TARUGO_") > 0 Then
            ok = ft.Select2(n > 0, 0)
            If ok Then n = n + 1
        End If
        Set ft = ft.GetNextFeature
    Loop
    If n = 0 Then Exit Sub
    Dim fo As SldWorks.Feature
    Set fo = md.FeatureManager.InsertFeatureTreeFolder2(swFeatureTreeFolder_Containing)
    If fo Is Nothing Then
        Anota "AVISO no se pudo crear la carpeta de tarugos (" & n & " cortes); quedan sueltos en el arbol"
    Else
        ' Nombre "TARUGOS MACRO". Dos carpetas no pueden llamarse igual: si la pieza ya tenia una
        ' (mecanizado anterior) SolidWorks rechaza el nombre y la carpeta se quedaba como "Carpeta".
        ' Se comprueba el nombre resultante y, si esta ocupado, se prueba con " 2", " 3"...
        Dim nombreCarpeta As String, intento As Long
        nombreCarpeta = "TARUGOS MACRO"
        intento = 1
        On Error Resume Next
        Do
            Err.Clear
            fo.Name = nombreCarpeta
            If Err.Number = 0 Then
                If fo.Name = nombreCarpeta Then Exit Do
            End If
            intento = intento + 1
            nombreCarpeta = "TARUGOS MACRO " & intento
        Loop While intento <= 20
        If fo.Name <> nombreCarpeta Then nombreCarpeta = fo.Name
        On Error GoTo 0
        Anota "Carpeta """ & nombreCarpeta & """ creada con " & n & " cortes"
        If Left(nombreCarpeta, 13) <> "TARUGOS MACRO" Then Anota "AVISO la carpeta no se pudo renombrar; renombrala a mano si quieres"
    End If
    md.ClearSelection2 True
    Exit Sub
fallo:
    Anota "AVISO error al crear la carpeta de tarugos: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    md.ClearSelection2 True
End Sub

' Elimina el ultimo croquis de la pieza (el que acaba de crear esta macro y no se uso)
Private Sub BorraUltimoCroquis(md As SldWorks.ModelDoc2)
    Dim sf As SldWorks.Feature
    Set sf = CroquisNuevo(md)
    If sf Is Nothing Then Exit Sub
    md.ClearSelection2 True
    sf.Select2 False, 0
    md.EditDelete
    md.ClearSelection2 True
End Sub

' Croquis que existian antes de abrir el nuevo. Asi se reconoce el croquis creado por la macro
' aunque no sea el ultimo del arbol (barra de retroceso, croquis reordenados o dentro de carpetas).
Private Sub SnapshotCroquis(md As SldWorks.ModelDoc2)
    Set croquisPrevios = New Collection
    Dim ft As SldWorks.Feature
    Set ft = md.FirstFeature
    On Error Resume Next
    Do While Not ft Is Nothing
        If ft.GetTypeName2 = "ProfileFeature" Then croquisPrevios.Add True, "k" & ft.Name
        Set ft = ft.GetNextFeature
    Loop
    On Error GoTo 0
End Sub

' El croquis (ProfileFeature) cuyo nombre no estaba en el snapshot; si no se encuentra, el ultimo.
Private Function CroquisNuevo(md As SldWorks.ModelDoc2) As SldWorks.Feature
    Dim ft As SldWorks.Feature, ultimo As SldWorks.Feature, nuevo As SldWorks.Feature, esta As Boolean
    Set ft = md.FirstFeature
    Do While Not ft Is Nothing
        If ft.GetTypeName2 = "ProfileFeature" Then
            Set ultimo = ft
            esta = False
            If Not croquisPrevios Is Nothing Then
                On Error Resume Next
                esta = croquisPrevios("k" & ft.Name)
                If Err.Number <> 0 Then esta = False
                Err.Clear
                On Error GoTo 0
            End If
            If Not esta Then Set nuevo = ft
        End If
        Set ft = ft.GetNextFeature
    Loop
    If nuevo Is Nothing Then Set CroquisNuevo = ultimo Else Set CroquisNuevo = nuevo
End Function

Private Function NombreCroquisNuevo(md As SldWorks.ModelDoc2) As String
    Dim f As SldWorks.Feature
    Set f = CroquisNuevo(md)
    If f Is Nothing Then NombreCroquisNuevo = "ninguno" Else NombreCroquisNuevo = f.Name
End Function

Private Function CortaCiego(md As SldWorks.ModelDoc2, ByVal prof As Double, ByVal invierteSentido As Boolean, Optional cuerpo As SldWorks.Body2 = Nothing) As SldWorks.Feature
    Dim sf As SldWorks.Feature
    Set sf = CroquisNuevo(md)
    If sf Is Nothing Then Exit Function
    md.ClearSelection2 True
    sf.Select2 False, 0
    If cuerpo Is Nothing Then
        ' ensamblaje / pieza de un solo cuerpo: como siempre (autoseleccion de cuerpos)
        Set CortaCiego = md.FeatureManager.FeatureCut4(True, False, invierteSentido, 0, 0, prof, 0, _
            False, False, False, False, 0, 0, False, False, False, False, False, True, True, _
            False, False, False, 0, 0, False, False)
    Else
        ' pieza multicuerpo: el corte se limita al cuerpo de la pieza (alcance de operacion):
        ' el cuerpo se selecciona con la marca 8 y UseAutoSelect = False. Con la autoseleccion el
        ' corte atravesaba tambien los otros cuerpos que hubiera detras (comprobado con una
        ' pieza de prueba). Con otras marcas (0, 1, 2, 4) FeatureCut4 devuelve Nothing.
        Dim sd As SldWorks.SelectData
        Set sd = md.SelectionManager.CreateSelectData
        sd.Mark = 8
        cuerpo.Select2 True, sd
        Set CortaCiego = md.FeatureManager.FeatureCut4(True, False, invierteSentido, 0, 0, prof, 0, _
            False, False, False, False, 0, 0, False, False, False, False, False, True, False, _
            False, False, False, 0, 0, False, False)
    End If
End Function
