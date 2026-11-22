# Tarugos Pro - contexto para continuar el desarrollo

Macro VBA de SolidWorks (probada en 2024) que detecta uniones entre tableros de un ensamblaje, propone tarugos,
deja revisarlos y crea taladros reales (cortes extruidos) en los archivos de cada pieza. Ver `LEEME.md` para el uso.

## Estructura
- `TarugosPro.swp`: proyecto compilado con TODO ya cargado (es lo que se ejecuta). Entrada: `main` -> `Tarugos_Principal`.
- `src/`: modulos .bas/.cls (fuente de verdad, ultima version). `formularios/`: codigo de los 2 UserForms.
- Los .frm NO se pueden importar en este SolidWorks ("memoria insuficiente"): los controles se crean por codigo.
  Para reconstruir el .swp: Insertar > UserForm, renombrar (`frmPrincipal`, `frmRevision`), pegar el codigo del .txt, e importar los .bas.
- Todos los modulos con procedimientos publicos llevan `Option Private Module` (para que solo salga `main` en Macro > Ejecutar).

## Como trabajar con SolidWorks desde Claude (lecciones aprendidas)
- No hay API COM desde fuera: se maneja la interfaz (computer use). Editor: Herramientas > Macro > Editar > abrir el .swp;
  quitar un modulo (clic derecho > Quitar, responder NO a exportar) e importarlo con Ctrl+M. F5 ejecuta; Depuracion > Compilar comprueba.
- **No llamar a open_application para traer SolidWorks al frente: lanza una sesion nueva vacia.** Si se pierde el foco, esperar o pedir al usuario.
- **Permiso para controlar el editor VBA**: pedir `request_access` con exactamente `sldworks.exe` (el editor es una ventana del mismo proceso; "Microsoft Visual Basic..." da notInstalled). Con eso se puede abrir el editor (Herramientas > Macro > Editar), quitar un modulo (clic derecho > Quitar, No a exportar), importar con Ctrl+M (escribir la ruta del .bas en Nombre, sin saltos de linea, y Enter aparte), Depuracion > Compilar y GUARDAR con el icono del disquete de la barra del editor (Ctrl+S no siempre guarda); Alt+Q = Cerrar y volver. El .swp cambia de tamano al cerrar el editor (SolidWorks lo reescribe): hacer el commit DESPUES de cerrar.
- Tambien se puede inspeccionar SolidWorks por COM desde Python (pywin32, sin `-I`): `w.Dispatch(pythoncom.GetActiveObject("SldWorks.Application").QueryInterface(pythoncom.IID_IDispatch))`; los metodos sin argumentos se leen como propiedades, y `SelectByID2` necesita `VARIANT(VT_DISPATCH, None)` en el callout.
- Comprobar siempre con la tipo-libreria: `sldworks.tlb`/`swconst.tlb` (en la carpeta de instalacion). Hay un volcador en la sesion anterior; se puede reescribir con ITypeLib.
- Probar siempre sobre copias o ensamblajes de prueba (`mPrueba`: CreaCocina, CreaEnsamblajePrueba, etc., en una carpeta `Pruebas`).

## Trampas de la API ya resueltas (no repetir)
- `Surface.IsPlane` devuelve valores inconsistentes: usar `Surface.Identity = 4001`. Curvas: `Identity = 3001` recta, `3002` circulo.
- `GetBodies3` devuelve geometria en espacio de PIEZA salvo transformacion identidad; transformar una copia con `Body2.ApplyTransform`.
- `MathUtility.CreatePoint/CreateVector` con `Array(...)` falla: usar `Double()` tipado.
- Las caras se liberan al salir de la funcion: mantener los cuerpos vivos (coleccion `Keep`).
- `Display3` de cuerpos temporales: bucle infinito de repintado si no se protege (flag + minimo 0,3 s). Dibuja en el espacio del componente de referencia: compensar con su transformacion inversa.
- `FeatureCut4(Flip...)`: Flip = cortar por fuera; el sentido es `Dir`. `AddToDB = True` al crear el circulo para evitar inferencias.
- `Feature.CustomPropertyManager.Add3` falla: la marca de la macro se guarda en `Feature.Description` ("TARUGO_META=...").
- Las operaciones de corte son tipo `"ICE"` (no `"Cut"`).
- `FileCopy` falla con archivos abiertos: copiar con `Extension.SaveAs` + `swSaveAsOptions_Copy`.
- `AddComponent5` coloca el CENTRO de la caja envolvente en (X,Y,Z).
- `Face2.Select4` de caras de `GetBodies2` puede seleccionar otra cara: usar `SelectByID2` por posicion y verificar normal.

## Estado
Hecho: fases 1-8 (geometria, uniones, reparto, revision 3D, taladros reales, copias independientes, duplicados, deshacer, rendimiento, interfaz, CSV).
Probado: tableros rectangulares a 90 grados, hasta 90 componentes (0,5 s), SolidWorks 2024.
Pendiente / sin confirmar:
- Mecanizado SIN ventana por pieza: IMPOSIBLE (probado por COM sobre una copia). Oculta y sin activar: SelectByID2 no da la cara. Oculta + `ActivateDoc3`: funciona pero SolidWorks la hace visible. Activar y luego `Visible = False`: vuelve a activarse el ensamblaje. Solo el documento activo (con ventana) permite seleccionar y croquizar. Se abre + `ActivateDoc3` + taladro + guardado + `CloseDoc` (solo si la ventana no la tenia abierta el usuario). Taladrado confirmado el 22/11/2026.
- Revision 3D (mPreview): SolidWorks IGNORA el color de los cuerpos temporales en alambre (todos salen en el azul palido del modelo); solo la SELECCION se pinta con color vivo. Por eso: (1) al marcar un tarugo se seleccionan sus 2 piezas (Component2.Select2/DeSelect, solo las que puso la macro) y su ARO (Body2.Select2); (2) durante la revision el color "Elemento seleccionado 1" (`swSystemColorsSelectedItem1`=104) se pone en ROJO y se restaura en `Limpia`. El aro va en coordenadas de ensamblaje y con `Display3 Nothing`: compensado con invCB el resalte de la seleccion salia desplazado (fuera del modelo). `Redibuja` llama dos veces a `ActualizaResalte` (la 2.a, tras dibujar, selecciona el aro). Probado a mano desde la LISTA: funciona. El clic en el 3D sobre los tarugos NO es fiable (el clic cae en la geometria de la pieza): la lista es la forma de elegir.
- Probado solo en la fase de revision (lista, aro rojo, piezas resaltadas); sin probar en ejecucion real: gestion de errores en `Mecaniza`, comprobacion del ensamblaje analizado (`mMain.AsmAnalizado`), copia del ensamblaje, aviso de taladros cercanos, regla de huecos (`mHuecos`: mueve o descarta tarugos que caen en un hueco; solo caras con agujeros o contorno no convexo) y uniones cortas con 3/2/1 tarugos (`mPosiciones.CabenN`, `NotaReparto`). Hacer primero una ejecucion normal y comprobar 0 avisos / 0 errores.
- Revision con CASILLAS (frmRevision, probado el 22/11/2026 con el ensamblaje de `pruebas/`): lista con casilla por fila (`ListStyle=1`, `MultiSelect=1`); marcar = seleccionar (resalta en el 3D). Botones QUITAR / RESTAURAR MARCADOS (la accion limpia las marcas y conserva el scroll), MARCAR TODOS / DESMARCAR TODOS / MARCAR ELIMINADOS, contador "Marcados: n de N". Un clic en el 3D sobre un tarugo lo AGREGA a los marcados y un clic que no toca ningun tarugo NO borra las marcas. Con mas de 25 marcados (`MAX_RESALTE`) solo se resaltan las piezas (sin aro). RESTAURAR no reactiva los "sin material". El codigo del formulario vive en `formularios/frmRevision_codigo.txt`: se pega en el editor (clic en el codigo, Ctrl+A, Ctrl+V; el portapapeles se llena con `Get-Content -Raw ... | Set-Clipboard`).
- Quedan en mTaladros `Diagnostica` y su llamada (codigo temporal): quitar cuando se de por cerrado el taladrado.
- Piezas virtuales (simetrias) se omiten; piezas inclinadas, canto con canto, contornos curvos: sin probar.
- Analisis "completo" del documento original: no implementado. Compatibilidad SW 2025/2026 y PDM: sin probar / no integrado
  (todo el acceso a archivos esta en `mGuardado` para anadirlo).
- Informe CSV (`mInformes`): escrito, no verificado con una ejecucion real.
