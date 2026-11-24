Attribute VB_Name = "mPrueba"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mPrueba - Crea un ensamblaje de PRUEBA con dos tableros reales (archivos .SLDPRT)
' en C:\Claude\Macros SolidWorks\Pruebas, para probar el taladrado sin tocar
' ninguno de tus diseños.
'   Prueba_Tapa.SLDPRT     1000 x 300 x 19 mm (tumbada)
'   Prueba_Costado.SLDPRT    19 x 300 x 600 mm (de pie, apoyado sobre la tapa)
' Ejecutar  CreaEnsamblajePrueba  y despues  Fase3_Revisar.
'==============================================================================

Private Const CARPETA As String = "C:\Claude\Macros SolidWorks\Pruebas\"

Public Sub CreaEnsamblajePrueba()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    On Error Resume Next
    MkDir CARPETA
    On Error GoTo 0

    Dim pTapa As String, pCost As String, pAsm As String
    pTapa = CARPETA & "Prueba_Tapa.SLDPRT"
    pCost = CARPETA & "Prueba_Costado.SLDPRT"
    pAsm = CARPETA & "Prueba_Tarugos.SLDASM"

    ' si hay documentos de prueba abiertos de una ejecucion anterior, se cierran
    Dim d As SldWorks.ModelDoc2, rr As Variant
    For Each rr In Array(pAsm, pTapa, pCost)
        Set d = swApp.GetOpenDocumentByName(CStr(rr))
        If Not d Is Nothing Then swApp.CloseDoc d.GetTitle
    Next rr

    If Dir(pTapa) <> "" Or Dir(pCost) <> "" Or Dir(pAsm) <> "" Then
        If MsgBox("Ya existen archivos de prueba en " & CARPETA & vbCrLf & "Se sobrescribiran. ¿Continuar?", _
                  vbYesNo + vbQuestion, "Tarugos - prueba") <> vbYes Then Exit Sub
    End If

    If Not CreaCaja(swApp, pTapa, 1#, 0.3, 0.019) Then MsgBox "No se pudo crear la tapa": Exit Sub
    If Not CreaCaja(swApp, pCost, 0.019, 0.3, 0.6) Then MsgBox "No se pudo crear el costado": Exit Sub

    ' AddComponent5 necesita los documentos cargados en memoria
    Dim e0 As Long, w0 As Long
    swApp.OpenDoc6 pTapa, swDocPART, swOpenDocOptions_Silent, "", e0, w0
    swApp.OpenDoc6 pCost, swDocPART, swOpenDocOptions_Silent, "", e0, w0

    Dim tpl As String
    tpl = swApp.GetUserPreferenceStringValue(swDefaultTemplateAssembly)
    Dim md As SldWorks.ModelDoc2
    Set md = swApp.NewDocument(tpl, 0, 0, 0)
    If md Is Nothing Then MsgBox "No se pudo crear el ensamblaje": Exit Sub
    Dim asm As SldWorks.AssemblyDoc
    Set asm = md
    Dim c1 As SldWorks.Component2, c2 As SldWorks.Component2
    ' OJO: AddComponent5 coloca el CENTRO de la caja envolvente de la pieza en (X,Y,Z)
    Set c1 = asm.AddComponent5(pTapa, 0, "", False, "", 0.5, 0.15, 0.0095)
    Set c2 = asm.AddComponent5(pCost, 0, "", False, "", 0.0095, 0.15, 0.319)   ' costado apoyado sobre la tapa (z = 19 mm)
    If c1 Is Nothing Or c2 Is Nothing Then MsgBox "No se pudieron insertar los componentes.", vbExclamation, "Tarugos - prueba": Exit Sub
    Dim e As Long, w As Long
    md.Extension.SaveAs pAsm, 0, 1, Nothing, e, w
    md.ViewZoomtofit2
    MsgBox "Ensamblaje de prueba creado:" & vbCrLf & pAsm, vbInformation, "Tarugos - prueba"
End Sub

'------------------------------------------------------------------------------
' Prueba 2 (Fase 5): UN mismo archivo de costado usado DOS veces, con taladros
' distintos en cada instancia (una sobre la tapa de 300 mm, otra sobre una base de
' 200 mm). Sirve para probar el aviso de pieza compartida y las copias independientes.
'------------------------------------------------------------------------------
Public Sub CreaEnsamblajePrueba2()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    On Error Resume Next
    MkDir CARPETA
    On Error GoTo 0

    Dim pTapa As String, pCost As String, pBase As String, pAsm As String
    pTapa = CARPETA & "Prueba_Tapa.SLDPRT"
    pCost = CARPETA & "Prueba_Costado.SLDPRT"
    pBase = CARPETA & "Prueba_Base2.SLDPRT"
    pAsm = CARPETA & "Prueba_Compartida.SLDASM"

    Dim d As SldWorks.ModelDoc2, rr As Variant
    For Each rr In Array(pAsm, pTapa, pCost, pBase, CARPETA & "Prueba_Tarugos.SLDASM")
        Set d = swApp.GetOpenDocumentByName(CStr(rr))
        If Not d Is Nothing Then swApp.CloseDoc d.GetTitle
    Next rr
    If Dir(pAsm) <> "" Or Dir(pBase) <> "" Then
        If MsgBox("Se sobrescribiran los archivos de prueba (tapa, costado, base2 y ensamblaje). ¿Continuar?", _
                  vbYesNo + vbQuestion, "Tarugos - prueba 2") <> vbYes Then Exit Sub
    End If

    If Not CreaCaja(swApp, pTapa, 1#, 0.3, 0.019) Then MsgBox "No se pudo crear la tapa": Exit Sub
    If Not CreaCaja(swApp, pCost, 0.019, 0.3, 0.6) Then MsgBox "No se pudo crear el costado": Exit Sub
    If Not CreaCaja(swApp, pBase, 0.4, 0.2, 0.019) Then MsgBox "No se pudo crear la base2": Exit Sub

    Dim e0 As Long, w0 As Long
    swApp.OpenDoc6 pTapa, swDocPART, swOpenDocOptions_Silent, "", e0, w0
    swApp.OpenDoc6 pCost, swDocPART, swOpenDocOptions_Silent, "", e0, w0
    swApp.OpenDoc6 pBase, swDocPART, swOpenDocOptions_Silent, "", e0, w0

    Dim tpl As String
    tpl = swApp.GetUserPreferenceStringValue(swDefaultTemplateAssembly)
    Dim md As SldWorks.ModelDoc2
    Set md = swApp.NewDocument(tpl, 0, 0, 0)
    If md Is Nothing Then MsgBox "No se pudo crear el ensamblaje": Exit Sub
    Dim asm As SldWorks.AssemblyDoc
    Set asm = md
    ' AddComponent5 coloca el CENTRO de la caja envolvente de la pieza en (X,Y,Z)
    Dim c1 As SldWorks.Component2, c2 As SldWorks.Component2, c3 As SldWorks.Component2, c4 As SldWorks.Component2
    Set c1 = asm.AddComponent5(pTapa, 0, "", False, "", 0.5, 0.15, 0.0095)
    Set c2 = asm.AddComponent5(pCost, 0, "", False, "", 0.0095, 0.15, 0.319)
    Set c3 = asm.AddComponent5(pBase, 0, "", False, "", 2.2, 0.1, 0.0095)
    Set c4 = asm.AddComponent5(pCost, 0, "", False, "", 2.0095, 0.15, 0.319)     ' 2a instancia del MISMO costado
    If c1 Is Nothing Or c2 Is Nothing Or c3 Is Nothing Or c4 Is Nothing Then
        MsgBox "No se pudieron insertar los componentes.", vbExclamation, "Tarugos - prueba 2": Exit Sub
    End If
    Dim e As Long, w As Long
    md.Extension.SaveAs pAsm, 0, 1, Nothing, e, w
    md.ViewZoomtofit2
    MsgBox "Ensamblaje de prueba 2 creado:" & vbCrLf & pAsm, vbInformation, "Tarugos - prueba 2"
End Sub

'------------------------------------------------------------------------------
' Prueba de RENDIMIENTO (Fase 7): N modulos (tapa + 2 costados) = 3N componentes,
' reutilizando los archivos Prueba_Tapa y Prueba_Costado (pieza repetida muchas veces).
'------------------------------------------------------------------------------
Public Sub CreaEnsamblajeGrande()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    Dim n As Long
    n = 30
    Dim pTapa As String, pCost As String, pAsm As String
    pTapa = CARPETA & "Prueba_Tapa.SLDPRT"
    pCost = CARPETA & "Prueba_Costado.SLDPRT"
    pAsm = CARPETA & "Prueba_Grande.SLDASM"
    If Dir(pTapa) = "" Or Dir(pCost) = "" Then MsgBox "Ejecuta antes CreaEnsamblajePrueba.", vbExclamation: Exit Sub
    Dim e0 As Long, w0 As Long
    swApp.OpenDoc6 pTapa, swDocPART, swOpenDocOptions_Silent, "", e0, w0
    swApp.OpenDoc6 pCost, swDocPART, swOpenDocOptions_Silent, "", e0, w0
    Dim tpl As String
    tpl = swApp.GetUserPreferenceStringValue(swDefaultTemplateAssembly)
    Dim md As SldWorks.ModelDoc2
    Set md = swApp.NewDocument(tpl, 0, 0, 0)
    Dim asm As SldWorks.AssemblyDoc
    Set asm = md
    Dim k As Long, xo As Double, yo As Double, c As SldWorks.Component2
    For k = 0 To n - 1
        xo = (k Mod 10) * 1.2
        yo = (k \ 10) * 0.5
        Set c = asm.AddComponent5(pTapa, 0, "", False, "", xo + 0.5, yo + 0.15, 0.0095)
        Set c = asm.AddComponent5(pCost, 0, "", False, "", xo + 0.0095, yo + 0.15, 0.319)
        Set c = asm.AddComponent5(pCost, 0, "", False, "", xo + 0.9905, yo + 0.15, 0.319)
    Next k
    Dim e As Long, w As Long
    md.Extension.SaveAs pAsm, 0, 1, Nothing, e, w
    md.ViewZoomtofit2
    MsgBox "Ensamblaje grande creado con " & (3 * n) & " componentes:" & vbCrLf & pAsm, vbInformation, "Tarugos - rendimiento"
End Sub

'------------------------------------------------------------------------------
' Mueble de cocina de ejemplo (600 x 720 x 560 mm, tablero de 18 mm):
' 2 costados, base, trasera y puerta, como archivos externos en la carpeta Pruebas.
' Ejes: X ancho, Y fondo (0 = frente), Z altura.
'------------------------------------------------------------------------------
Public Sub CreaCocina()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    On Error Resume Next
    MkDir CARPETA
    On Error GoTo 0

    Dim nom As Variant, rutas(0 To 4) As String, i As Long
    nom = Array("Cocina_CostadoIzq", "Cocina_CostadoDer", "Cocina_Base", "Cocina_Trasera", "Cocina_Puerta")
    For i = 0 To 4: rutas(i) = CARPETA & nom(i) & ".SLDPRT": Next i
    Dim pAsm As String
    pAsm = CARPETA & "Cocina.SLDASM"

    Dim d As SldWorks.ModelDoc2
    For i = 0 To 4
        Set d = swApp.GetOpenDocumentByName(rutas(i))
        If Not d Is Nothing Then swApp.CloseDoc d.GetTitle
    Next i
    Set d = swApp.GetOpenDocumentByName(pAsm)
    If Not d Is Nothing Then swApp.CloseDoc d.GetTitle

    If Not CreaCaja(swApp, rutas(0), 0.018, 0.56, 0.72) Then MsgBox "Error costado izq": Exit Sub
    If Not CreaCaja(swApp, rutas(1), 0.018, 0.56, 0.72) Then MsgBox "Error costado der": Exit Sub
    If Not CreaCaja(swApp, rutas(2), 0.564, 0.56, 0.018) Then MsgBox "Error base": Exit Sub
    If Not CreaCaja(swApp, rutas(3), 0.564, 0.018, 0.702) Then MsgBox "Error trasera": Exit Sub
    If Not CreaCaja(swApp, rutas(4), 0.596, 0.018, 0.716) Then MsgBox "Error puerta": Exit Sub

    Dim e0 As Long, w0 As Long
    For i = 0 To 4
        swApp.OpenDoc6 rutas(i), swDocPART, swOpenDocOptions_Silent, "", e0, w0
    Next i

    Dim tpl As String
    tpl = swApp.GetUserPreferenceStringValue(swDefaultTemplateAssembly)
    Dim md As SldWorks.ModelDoc2
    Set md = swApp.NewDocument(tpl, 0, 0, 0)
    If md Is Nothing Then MsgBox "No se pudo crear el ensamblaje": Exit Sub
    Dim asm As SldWorks.AssemblyDoc
    Set asm = md
    ' AddComponent5 coloca el CENTRO de la caja envolvente en (X,Y,Z), en metros
    Dim c As SldWorks.Component2
    Set c = asm.AddComponent5(rutas(0), 0, "", False, "", 0.009, 0.28, 0.36)
    Set c = asm.AddComponent5(rutas(1), 0, "", False, "", 0.591, 0.28, 0.36)
    Set c = asm.AddComponent5(rutas(2), 0, "", False, "", 0.3, 0.28, 0.009)
    Set c = asm.AddComponent5(rutas(3), 0, "", False, "", 0.3, 0.551, 0.369)
    Set c = asm.AddComponent5(rutas(4), 0, "", False, "", 0.3, -0.011, 0.36)
    Dim e As Long, w As Long
    md.Extension.SaveAs pAsm, 0, 1, Nothing, e, w
    md.ViewZoomtofit2
    MsgBox "Cocina creada:" & vbCrLf & pAsm, vbInformation, "Tarugos - cocina"
End Sub

Private Function CreaCaja(swApp As SldWorks.SldWorks, ruta As String, ByVal x As Double, ByVal y As Double, ByVal z As Double) As Boolean
    Dim tpl As String
    tpl = swApp.GetUserPreferenceStringValue(swDefaultTemplatePart)
    Dim pm As SldWorks.ModelDoc2
    Set pm = swApp.NewDocument(tpl, 0, 0, 0)
    If pm Is Nothing Then Exit Function

    Dim feat As SldWorks.Feature, plano As SldWorks.Feature, sk As SldWorks.Feature
    Set feat = pm.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "RefPlane" And plano Is Nothing Then Set plano = feat   ' 1er plano = Alzado/Front
        Set feat = feat.GetNextFeature
    Loop
    plano.Select2 False, 0
    pm.SketchManager.InsertSketch True
    pm.SketchManager.CreateCornerRectangle 0, 0, 0, x, y, 0
    pm.SketchManager.InsertSketch True

    Set feat = pm.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "ProfileFeature" Then Set sk = feat
        Set feat = feat.GetNextFeature
    Loop
    sk.Select2 False, 0

    Dim ext As SldWorks.Feature
    Set ext = pm.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, z, 0, _
              False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    If ext Is Nothing Then swApp.CloseDoc pm.GetTitle: Exit Function

    Dim e As Long, w As Long, ok As Boolean
    ok = pm.Extension.SaveAs(ruta, 0, 1, Nothing, e, w)
    swApp.CloseDoc pm.GetTitle
    CreaCaja = (Dir(ruta) <> "")
End Function
