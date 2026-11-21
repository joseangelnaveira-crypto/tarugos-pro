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
- Mecanizado SIN ventana por pieza (`DocumentVisible False`, sin `ActivateDoc3`): PROBADO Y FALLA (0 de 118: SelectByID2 no da la cara e InsertSketch no abre croquis). Se volvio a abrir + `ActivateDoc3` en `MecanizaDoc`; pendiente de confirmar que ahora si taladra.
- Tarugo seleccionado en rojo y mas grande: cargado, no probado visualmente. El clic directo sobre cilindros en 3D no es fiable.
- Piezas virtuales (simetrias) se omiten; piezas inclinadas, canto con canto, contornos curvos: sin probar.
- Analisis "completo" del documento original: no implementado. Compatibilidad SW 2025/2026 y PDM: sin probar / no integrado
  (todo el acceso a archivos esta en `mGuardado` para anadirlo).
- Informe CSV (`mInformes`): escrito, no verificado con una ejecucion real.
