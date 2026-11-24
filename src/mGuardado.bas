Attribute VB_Name = "mGuardado"
Option Explicit
Option Private Module   ' sus procedimientos no aparecen en la lista Macro > Ejecutar
'==============================================================================
' mGuardado - FASE 5: TODO el acceso a archivos pasa por este modulo.
'
' Motivo: hoy se trabaja con disco local, pero mas adelante habra que usar PDM
' (extraer antes de modificar, insertar despues). Si todas las copias, guardados
' y comprobaciones de solo-lectura estan aqui, ese cambio se hace en un solo sitio
' sin tocar el resto de la macro.
'==============================================================================

Public Const META_PROP As String = "TARUGO_META"   ' propiedad personalizada de cada operacion

' True si el documento no se puede guardar (abierto en solo lectura / sin extraer).
' PDM (futuro): aqui se comprobaria tambien que el archivo esta extraido por el usuario.
Public Function EsSoloLectura(md As SldWorks.ModelDoc2) As Boolean
    On Error Resume Next
    EsSoloLectura = CBool(md.IsOpenedReadOnly)
    If Dir(md.GetPathName, vbReadOnly) <> "" Then
        If (GetAttr(md.GetPathName) And vbReadOnly) <> 0 Then EsSoloLectura = True
    End If
End Function

' Carpeta de respaldo junto al ensamblaje. Devuelve la ruta (ya creada).
Public Function CarpetaRespaldo(amd As SldWorks.ModelDoc2) As String
    Dim c As String
    c = Left(amd.GetPathName, InStrRev(amd.GetPathName, "\")) & "Respaldo_Tarugos_" & Format(Now, "yyyymmdd_hhnnss")
    On Error Resume Next
    MkDir c
    On Error GoTo 0
    ' si la carpeta no existe (sin permisos, disco lleno...) se devuelve "": NO se debe mecanizar sin copia
    If Dir(c, vbDirectory) = "" Then CarpetaRespaldo = "" Else CarpetaRespaldo = c
End Function

' Copia el documento (aunque este abierto) con "guardar copia". Devuelve True si el
' archivo destino existe. No cambia el nombre ni el estado del documento abierto.
' (FileCopy falla con "permiso denegado" con los archivos abiertos en SolidWorks.)
Public Function CopiaA(md As SldWorks.ModelDoc2, ByVal destino As String) As Boolean
    Dim e As Long, w As Long, ok As Boolean
    ok = md.Extension.SaveAs(destino, 0, swSaveAsOptions_Silent + swSaveAsOptions_Copy, Nothing, e, w)
    CopiaA = (Dir(destino) <> "")
End Function

' Ruta libre para una copia independiente: <nombre>_T<n>.SLDPRT junto al original
Public Function RutaCopiaLibre(ByVal rutaOriginal As String) As String
    Dim baseN As String, n As Long, r As String
    baseN = Left(rutaOriginal, InStrRev(rutaOriginal, ".") - 1)
    n = 2
    Do
        r = baseN & "_T" & n & ".SLDPRT"
        If Dir(r) = "" Then Exit Do
        n = n + 1
    Loop
    RutaCopiaLibre = r
End Function

' Guarda UNA vez el documento. Devuelve el codigo de error de SolidWorks (0 = correcto).
Public Function GuardaDoc(md As SldWorks.ModelDoc2) As Long
    Dim e As Long, w As Long
    md.Save3 swSaveAsOptions_Silent, e, w
    GuardaDoc = e
End Function

' Sustituye (solo la instancia indicada) un componente por otro archivo.
Public Function ReemplazaInstancia(asm As SldWorks.AssemblyDoc, md As SldWorks.ModelDoc2, _
                                   comp As SldWorks.Component2, ByVal nuevaRuta As String) As Boolean
    md.ClearSelection2 True
    comp.Select4 False, Nothing, False
    ReemplazaInstancia = asm.ReplaceComponents2(nuevaRuta, "", False, swReplaceComponentsConfiguration_MatchName, True)
    md.ClearSelection2 True
End Function
