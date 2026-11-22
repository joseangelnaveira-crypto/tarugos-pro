Attribute VB_Name = "mPreview"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mPreview - FASE 3: previsualizacion 3D temporal de los tarugos.
'
' Cada tarugo se muestra como un cilindro TEMPORAL (Modeler.CreateBodyFromCyl +
' Body2.Display3 con swTempBodySelectable). No modifica ninguna pieza.
' SolidWorks borra los cuerpos temporales en cada repintado, asi que la clase
' cEventos llama a DibujaTodo desde RepaintPostNotify.
'
' Colores:  verde = activo | naranja = union dudosa forzada | gris = eliminado
'           rojo (mas grande y con aro) = seleccionado
'           rojo oscuro = sin material (cae en un hueco; descartado, ver mHuecos)
'==============================================================================

Public Type TFlat
    U As Long                ' indice de union
    K As Long                ' indice de tarugo dentro de la union
    Cuerpo As SldWorks.Body2
    Sel As Boolean
    Grande As Boolean       ' el cuerpo temporal actual esta dibujado en tamano 'seleccionado'
    Aro As SldWorks.Body2   ' disco fino y ancho alrededor del tarugo seleccionado (se crea al seleccionarlo)
End Type

Public Flat() As TFlat
Public NFlat As Long

Public swAppP As SldWorks.SldWorks
Public mdlP As SldWorks.ModelDoc2
Public Eventos As cEventos
Public Formulario As frmRevision

Private compBase As SldWorks.Component2
Private modeler As SldWorks.Modeler
Private invCB As SldWorks.MathTransform     ' inversa de la transformacion del componente de referencia
Const FACTOR_NORMAL As Double = 1.6     ' los cilindros se dibujan mas gruesos que el taladro real
Const FACTOR_SEL As Double = 3.5        ' el tarugo seleccionado se agranda para localizarlo
' En alambre el color casi no se aprecia, asi que el seleccionado lleva ademas un ARO: un disco
' fino y ancho (en alambre se ve como dos circulos grandes) centrado en el tarugo.
Const RADIO_ARO As Double = 0.06        ' m (60 mm de radio)
Const ESPESOR_ARO As Double = 0.001     ' m
Private modoOriginal As Long
Private modoCambiado As Boolean
Private dibujando As Boolean           ' proteccion contra reentrada en DibujaTodo
Private ultimoDibujo As Double     ' componente de referencia para Display3 (si hace falta)
Private errorDisplay As String
Public Llamadas As Long                     ' numero de repintados atendidos (diagnostico)
' Resalte de las dos piezas del tarugo seleccionado (se seleccionan como componentes: SolidWorks
' las pinta con su color de seleccion, tambien en alambre). Poner False para desactivarlo.
Private Const RESALTA_PIEZAS As Boolean = True
Private resNombres As String                ' piezas resaltadas ahora, entre barras: |Balda-1|Costado_IZQ-1|
Private seleccionPropia As Boolean          ' la macro esta cambiando la seleccion (SincronizaSeleccion lo ignora)
Private permitidos As Long                  ' dibujados que DibujaTodo concede sin esperar los 0,3 s (los pone Redibuja)

Private Function RGBl(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    RGBl = r + g * 256 + b * 65536
End Function

'------------------------------------------------------------------------------
' Crea los cuerpos temporales (uno por tarugo calculado)
'------------------------------------------------------------------------------
Public Sub ConstruyePreview(app As SldWorks.SldWorks, md As SldWorks.ModelDoc2)
    Set swAppP = app
    Set mdlP = md
    NFlat = 0
    Erase Flat
    resNombres = ""

    Set modeler = app.GetModeler

    ' Display3 dibuja el cuerpo temporal en el espacio del COMPONENTE de referencia (le aplica su
    ' transformacion). Si no se compensa, los cilindros quedan fuera de sitio o no se ven.
    Set compBase = Nothing: Set invCB = Nothing
    Dim ip As Long
    For ip = 0 To NPiezas - 1
        If Piezas(ip).Estado = "OK" Then
            Set compBase = Piezas(ip).Comp
            Exit For
        End If
    Next ip
    If Not compBase Is Nothing Then Set invCB = compBase.GetTotalTransform(False).Inverse

    Dim u As Long, k As Long
    For u = 0 To NUniones - 1
        For k = 0 To Uniones(u).NT - 1
            ReDim Preserve Flat(0 To NFlat)
            Flat(NFlat).U = u
            Flat(NFlat).K = k
            Set Flat(NFlat).Cuerpo = CreaCuerpo(u, k, FACTOR_NORMAL)
            If Not Flat(NFlat).Cuerpo Is Nothing Then NFlat = NFlat + 1
        Next k
    Next u

    ' Los tarugos quedan DENTRO de las piezas (no se verian en sombreado):
    ' durante la revision se pasa la vista a alambre y se restaura al cerrar.
    On Error Resume Next
    Dim vw As SldWorks.ModelView
    Set vw = md.GetFirstModelView
    If Not vw Is Nothing Then
        modoOriginal = vw.DisplayMode
        vw.DisplayMode = 1                  ' swViewDisplayMode_Wireframe
        modoCambiado = True
    End If
End Sub

' Cilindro temporal del tarugo k de la union u (radio = factor x radio del taladro)
Private Function CreaCuerpo(ByVal u As Long, ByVal k As Long, ByVal factor As Double) As SldWorks.Body2
    Dim cyl(0 To 7) As Double, base As TV3
    base = V3(Uniones(u).Pos(k).X - Uniones(u).N1.X * Cfg.Prof1 * MM, _
              Uniones(u).Pos(k).Y - Uniones(u).N1.Y * Cfg.Prof1 * MM, _
              Uniones(u).Pos(k).Z - Uniones(u).N1.Z * Cfg.Prof1 * MM)
    ' CylDimArray: centro de la base (3), eje (3), radio, altura. Debe ser Double().
    cyl(0) = base.X: cyl(1) = base.Y: cyl(2) = base.Z
    cyl(3) = Uniones(u).N1.X: cyl(4) = Uniones(u).N1.Y: cyl(5) = Uniones(u).N1.Z
    cyl(6) = Cfg.Diametro * MM / 2 * factor
    cyl(7) = (Cfg.Prof1 + Cfg.Prof2) * MM
    Dim bd As SldWorks.Body2
    Set bd = modeler.CreateBodyFromCyl(cyl)
    If Not bd Is Nothing And Not invCB Is Nothing Then bd.ApplyTransform invCB
    Set CreaCuerpo = bd
End Function

' Aro del tarugo k de la union u: disco centrado en el punto medio del tarugo, con su mismo eje
Private Function CreaAro(ByVal u As Long, ByVal k As Long) As SldWorks.Body2
    Dim cyl(0 To 7) As Double
    cyl(0) = Uniones(u).Pos(k).X - Uniones(u).N1.X * ESPESOR_ARO / 2
    cyl(1) = Uniones(u).Pos(k).Y - Uniones(u).N1.Y * ESPESOR_ARO / 2
    cyl(2) = Uniones(u).Pos(k).Z - Uniones(u).N1.Z * ESPESOR_ARO / 2
    cyl(3) = Uniones(u).N1.X: cyl(4) = Uniones(u).N1.Y: cyl(5) = Uniones(u).N1.Z
    cyl(6) = RADIO_ARO
    cyl(7) = ESPESOR_ARO
    Dim bd As SldWorks.Body2
    Set bd = modeler.CreateBodyFromCyl(cyl)
    If Not bd Is Nothing And Not invCB Is Nothing Then bd.ApplyTransform invCB
    Set CreaAro = bd
End Function

Private Function ColorDe(i As Long) As Long
    If Flat(i).Sel Then
        ColorDe = RGBl(230, 20, 20)            ' seleccionado (a quitar): rojo y mas grande
    ElseIf Uniones(Flat(i).U).SinMat(Flat(i).K) Then
        ColorDe = RGBl(150, 0, 0)              ' sin material (cae en un hueco): rojo oscuro
    ElseIf Not Uniones(Flat(i).U).Activo(Flat(i).K) Then
        ColorDe = RGBl(140, 140, 140)          ' ya eliminado: gris
    ElseIf Uniones(Flat(i).U).Estado = "VALIDA" Then
        ColorDe = RGBl(30, 190, 60)
    Else
        ColorDe = RGBl(255, 140, 0)
    End If
End Function

'------------------------------------------------------------------------------
' Dibuja todos los cuerpos temporales (se llama en cada repintado)
'------------------------------------------------------------------------------
' OJO: Display3 provoca un repintado, y RepaintPostNotify vuelve a llamar aqui.
' Sin proteccion se produce un bucle infinito que congela SolidWorks (ocurrio en
' las pruebas). Se evita con (1) un indicador de reentrada y (2) un minimo de
' tiempo entre dibujados.
' Excepcion: tras cambiar la seleccion (Redibuja) llegan repintados PROPIOS de SolidWorks
' (resalte de las piezas) despues del dibujado forzado; borran los cuerpos temporales y, con
' el minimo de 0,3 s, no se redibujaban (los tarugos desaparecian hasta el siguiente
' repintado). Por eso Redibuja concede unos pocos dibujados sin esperar (permitidos); son
' pocos y acotados, asi que no puede volver el bucle infinito.
Public Sub DibujaTodo()
    If dibujando Then Exit Sub
    If permitidos > 0 Then
        permitidos = permitidos - 1
    ElseIf ultimoDibujo <> 0 Then
        If Abs(Timer - ultimoDibujo) < 0.3 Then Exit Sub
    End If
    dibujando = True
    On Error Resume Next
    Llamadas = Llamadas + 1
    Dim i As Long
    For i = 0 To NFlat - 1
        Err.Clear
        If Flat(i).Sel <> Flat(i).Grande Then
            Dim nb As SldWorks.Body2
            Set nb = CreaCuerpo(Flat(i).U, Flat(i).K, IIf(Flat(i).Sel, FACTOR_SEL, FACTOR_NORMAL))
            If Not nb Is Nothing Then Set Flat(i).Cuerpo = nb: Flat(i).Grande = Flat(i).Sel
        End If
        Flat(i).Cuerpo.Display3 compBase, ColorDe(i), swTempBodySelectable
        If Flat(i).Sel Then
            If Flat(i).Aro Is Nothing Then Set Flat(i).Aro = CreaAro(Flat(i).U, Flat(i).K)
            If Not Flat(i).Aro Is Nothing Then Flat(i).Aro.Display3 compBase, ColorDe(i), swTempBodySelectable
        ElseIf Not Flat(i).Aro Is Nothing Then
            ' Al desmarcar hay que SOLTAR el aro: dejar de dibujarlo no basta, el cuerpo temporal
            ' sigue en pantalla mientras el objeto exista (igual que con el cilindro grande, que
            ' solo desaparece al sustituirlo por uno normal).
            Set Flat(i).Aro = Nothing
        End If
        If Err.Number <> 0 Then
            errorDisplay = Err.Number & ": " & Err.Description
            Err.Clear
        End If
    Next i
    ultimoDibujo = Timer
    dibujando = False
End Sub

Public Sub Redibuja()
    ' ActualizaResalte cambia la seleccion y eso provoca repintados propios: se hace ANTES de
    ' poner ultimoDibujo = 0. Si se hacia despues, uno de esos repintados consumia el permiso de
    ' dibujar (minimo 0,3 s entre dibujados) y el repintado final no redibujaba los tarugos:
    ' quedaba el aro del tarugo anterior y no aparecia el nuevo hasta hacer clic en el dibujo.
    Dim cambio As Boolean
    cambio = ActualizaResalte()
    ' Si se ha cambiado la seleccion, SolidWorks tarda en terminar sus repintados (resalte de las
    ' piezas) y cualquiera de ellos borra los cuerpos temporales. Se le deja medio segundo para
    ' que termine (DoEvents atiende los mensajes pendientes) y SOLO ENTONCES se fuerza el
    ' dibujado final, que asi es el ultimo.
    If cambio Then EsperaSolidWorks 0.5
    ultimoDibujo = 0                       ' forzar el dibujado en el repintado que sigue
    permitidos = 3                         ' y en los 2 que pueden llegar despues (ver DibujaTodo)
    If Not mdlP Is Nothing Then mdlP.GraphicsRedraw2
End Sub

Private Sub EsperaSolidWorks(ByVal seg As Double)
    Dim t As Double
    t = Timer
    Do While Abs(Timer - t) < seg
        DoEvents
    Loop
End Sub

' Lista de nombres entre barras ("|a|b|"): anade n si no estaba
Private Function AnadeNombre(ByVal lista As String, ByVal n As String) As String
    If lista = "" Then lista = "|"
    If InStr(lista, "|" & n & "|") = 0 Then lista = lista & n & "|"
    AnadeNombre = lista
End Function

'------------------------------------------------------------------------------
' Resalta las dos piezas (componentes) de cada tarugo seleccionado, para localizarlo
' mejor en alambre. Se ANADEN a la seleccion (no se borra nada: los tarugos que el
' usuario haya seleccionado con el raton siguen seleccionados) y al desmarcar se
' quitan SOLO las que puso la macro (resNombres). No modifica el ensamblaje.
'------------------------------------------------------------------------------
Private Function ActualizaResalte() As Boolean
    If Not RESALTA_PIEZAS Then Exit Function
    If mdlP Is Nothing Then Exit Function
    On Error Resume Next
    Dim i As Long, u As Long, deseadas As String
    For i = 0 To NFlat - 1
        If Flat(i).Sel Then
            u = Flat(i).U
            deseadas = AnadeNombre(deseadas, Piezas(Uniones(u).P1).Nombre)
            deseadas = AnadeNombre(deseadas, Piezas(Uniones(u).P2).Nombre)
        End If
    Next i
    If deseadas = "" And resNombres = "" Then Exit Function

    ' componentes seleccionados ahora mismo (por la macro o por el usuario)
    Dim sm As SldWorks.SelectionMgr, c As SldWorks.Component2, sel As String, n As Long
    Set sm = mdlP.SelectionManager
    n = sm.GetSelectedObjectCount2(-1)
    For i = 1 To n
        If sm.GetSelectedObjectType3(i, -1) = swSelCOMPONENTS Then
            Set c = sm.GetSelectedObject6(i, -1)
            If Not c Is Nothing Then sel = AnadeNombre(sel, c.Name2)
        End If
    Next i

    seleccionPropia = True
    Dim p As Long, nom As String
    For p = 0 To NPiezas - 1
        nom = "|" & Piezas(p).Nombre & "|"
        If InStr(deseadas, nom) > 0 Then
            If InStr(sel, nom) = 0 Then Piezas(p).Comp.Select2 True, 0: ActualizaResalte = True
        ElseIf InStr(resNombres, nom) > 0 Then
            If InStr(sel, nom) > 0 Then Piezas(p).Comp.DeSelect: ActualizaResalte = True
        End If
    Next p
    seleccionPropia = False
    resNombres = deseadas
End Function

Public Function ErrorDeDisplay() As String
    ErrorDeDisplay = errorDisplay
End Function

'------------------------------------------------------------------------------
' Sincroniza la seleccion hecha con el raton en el modelo con la lista del formulario
'------------------------------------------------------------------------------
Public Sub SincronizaSeleccion()
    If Formulario Is Nothing Then Exit Sub
    If seleccionPropia Then Exit Sub        ' la seleccion la esta cambiando la macro (ActualizaResalte)
    On Error Resume Next
    Dim sm As SldWorks.SelectionMgr
    Set sm = mdlP.SelectionManager
    Dim n As Long, i As Long, j As Long, o As Object
    n = sm.GetSelectedObjectCount2(-1)
    Dim hit() As Boolean
    ReDim hit(0 To NFlat)
    For i = 1 To n
        If sm.GetSelectedObjectType3(i, -1) = swSelSOLIDBODIES Then
            Set o = sm.GetSelectedObject6(i, -1)
            For j = 0 To NFlat - 1
                If swAppP.IsSame(o, Flat(j).Cuerpo) Then hit(j) = True
            Next j
        End If
    Next i
    ' Se sincroniza SIEMPRE, tambien cuando no hay ningun tarugo seleccionado en el modelo:
    ' si no, al desmarcar un tarugo en el 3D (clic en vacio o Ctrl+clic) quedaba marcado
    ' en la lista y dibujado en rojo con su aro.
    Formulario.SeleccionaDesdeModelo hit
End Sub

Public Sub Limpia()
    On Error Resume Next
    If Not Eventos Is Nothing Then Eventos.Desconecta
    Set Eventos = Nothing
    NFlat = 0
    Erase Flat
    Set Formulario = Nothing
    If modoCambiado Then
        Dim vw As SldWorks.ModelView
        Set vw = mdlP.GetFirstModelView
        If Not vw Is Nothing Then vw.DisplayMode = modoOriginal
        modoCambiado = False
    End If
    Redibuja
    Set mdlP = Nothing
End Sub
