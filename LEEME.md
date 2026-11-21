# Tarugos Pro - macro VBA para SolidWorks

Detecta uniones entre tableros de un ensamblaje, calcula la distribucion de tarugos y crea
**taladros reales** (cortes extruidos editables) en los archivos de cada pieza.
Probada en SolidWorks 2024 (revision 32.3). Las funciones de la API usadas existen tambien en
2025/2026, pero **no se ha probado en esas versiones**.

## Instalacion
**En otro equipo basta copiar UN archivo: `TarugosPro.swp`** (lleva dentro todos los modulos, la clase y los formularios).
Abrirlo con Herramientas > Macro > Editar... Lo de `src/` y `formularios/` solo hace falta para modificar el codigo o reconstruir el .swp.
Si se cambia el codigo de `src/`, hay que volver a guardar el .swp desde el editor de SolidWorks para que no quede desfasado.

1. SolidWorks: Herramientas > Macro > Editar... y abrir `TarugosPro.swp` (esta carpeta).
2. Para ejecutar: abre un ensamblaje y, en el editor, F5 > `Tarugos_Principal`
   (o Herramientas > Macro > Ejecutar... y asigna la macro a un boton si quieres).
3. El proyecto se reconstruye importando los modulos de la carpeta `Fase8` (.bas / .cls) y creando
   un UserForm llamado `frmPrincipal` con el codigo de `frmPrincipal_codigo.txt`, y otro
   `frmRevision` con `frmRevision_codigo.txt` (los .frm no se pueden importar: los controles se crean por codigo).

## Uso
1. `Tarugos_Principal` abre la ventana con los parametros. **AUTOMATICO** analiza todo el ensamblaje;
   **MANUAL**: selecciona antes 2 o mas componentes y pulsa MANUAL (solo las uniones entre ellos).
2. Aparece la **revision**: lista de tarugos y cilindros temporales en el modelo (se ven en vista de alambre).
   Puedes eliminar/restaurar un tarugo o una union completa. Las uniones dudosas empiezan desactivadas.
3. **ACEPTAR Y MECANIZAR** pide confirmacion, hace copia de seguridad de cada pieza
   (carpeta `Respaldo_Tarugos_<fecha>` junto al ensamblaje), crea los taladros y guarda cada archivo una vez.
4. **DESHACER MECANIZADO** (o `Fase6_Deshacer`) elimina solo las operaciones creadas por la macro.

## Parametros (milimetros)
| Parametro | Defecto | Significado |
|---|---|---|
| Diametro | 8 | Diametro del taladro/tarugo |
| Profundidad pieza 1 | 31 | Pieza cuyo CANTO se apoya en la otra |
| Profundidad pieza 2 | 11 | Pieza que recibe el canto, en su cara |
| Distancia a los extremos | 50 | Del primer/ultimo tarugo al borde de la union |
| Longitud de referencia | 1000 | Hasta esta longitud: numero inicial de tarugos |
| Numero inicial | 3 | |
| Incremento de longitud | 300 | Cada 300 mm completos por encima de la referencia... |
| Incremento de tarugos | 1 | ...se anaden estos tarugos |
| Tolerancia geometrica | 0,1 | Para decidir si dos caras estan en contacto |

N = N inicial + IncN x ENTERO((L - Lref) / IncLong). Los ajustes se guardan en el registro de Windows
(`HKCU\Software\VB and VBA Program Settings\TarugosPro`).

## Resultados
- Operaciones nombradas `UNION_001_TARUGO_002_<pieza>_D8_P31-11`, con su posicion en el comentario (`TARUGO_META=...`).
- Informes: `Informe_Fase5.txt` y `Informe_Tarugos.csv` (separador `;`) en la carpeta de respaldo;
  `Tarugos_Fase2_*.txt` junto al ensamblaje con el analisis y los tiempos.
- Volver a ejecutar no duplica: los taladros ya creados se reconocen y se omiten.

## Modificar o eliminar taladros
- Editar: en el arbol de la pieza, clic derecho en la operacion `UNION_...` > Editar operacion.
- Eliminar todos los de la macro: `Fase6_Deshacer`. No se tocan operaciones que no lleven la marca de la macro.

## Piezas repetidas
Si una misma pieza aparece varias veces y las instancias necesitan taladros distintos, la macro pregunta:
SI = modificar el original (afecta a todas), NO = copia independiente (`<pieza>_T2.SLDPRT`) sustituida en este
ensamblaje, CANCELAR = omitir.

## Limitaciones conocidas
- Probada solo con tableros rectangulares a 90 grados, pocas decenas de piezas y SolidWorks 2024.
- Piezas inclinadas, uniones canto con canto (quedan "dudosas"), contornos curvos y geometrias complejas: sin probar.
- Piezas **virtuales** (guardadas dentro del ensamblaje, p. ej. las creadas por simetria) se omiten: guardalas
  antes como archivo externo.
- Solo hay un modo de analisis (rapido, con cajas envolventes); el "analisis completo" del documento original no esta implementado.
- Los cilindros de previsualizacion quedan dentro de las piezas: hay que verlos en alambre; la seleccion con clic en 3D no esta confirmada.
- Deshacer solo borra: no restaura desde la copia de seguridad ni revierte la sustitucion por copias independientes.
- Cancelar con ESC durante el analisis esta implementado pero no probado.
- El informe CSV esta escrito pero no se ha comprobado su contenido en una ejecucion real.
- PDM: no integrado. Todo el acceso a archivos esta en `mGuardado` para anadirlo mas adelante.
- Los formularios no son modales para el usuario en la revision; si SolidWorks se queda congelado, cerrar el proceso.

## Modulos
mConfig (parametros) - mVec, mPoligono (geometria 2D/3D) - mGeometria (piezas y caras) - mUniones (deteccion) -
mPosiciones (reparto) - mPreview, mEdicion, cEventos, frmRevision (revision) - mTaladros (cortes reales) -
mGuardado (acceso a archivos) - mDuplicados (marcas y duplicados) - mRecuperacion (deshacer) - mOptim (progreso) -
mInformes (CSV) - mPrueba (ensamblajes de prueba) - mMain (puntos de entrada) - frmPrincipal.
