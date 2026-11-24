# Tarugos Pro - macro VBA para SolidWorks

Detecta uniones entre tableros, calcula la distribucion de tarugos, deja revisarlos y crea
**taladros reales** (cortes extruidos editables). Funciona con **ensamblajes** (un archivo por pieza) y con
**piezas multicuerpo** (cada cuerpo solido es un tablero), con el mismo formulario y el mismo flujo.
Probada en SolidWorks 2024 (SP3.1). Las funciones de la API usadas existen tambien en 2025/2026,
pero **no se ha probado en esas versiones**.

## Instalacion
**En otro equipo basta copiar UN archivo: `TarugosPro.swp`** (lleva dentro todos los modulos, la clase y los formularios).
`src/` y `formularios/` solo hacen falta para modificar el codigo o reconstruir el `.swp`.

### Ejecutar
- Herramientas > Macro > Ejecutar... y elegir `TarugosPro.swp` (procedimiento `main`).

### Boton en SolidWorks
1. Herramientas > Personalizar > pestana **Comandos** > categoria **Macro**.
2. Arrastrar **Macro personalizada** a una barra de herramientas.
3. En el cuadro: **Macro** = Examinar > `TarugosPro.swp`; **Metodo** = `main` (`TarugosPro1.main`); elegir imagen y sugerencia.
4. Aceptar. (Si se mueve el `.swp` de carpeta hay que reasignar el boton. Tambien se puede asignar un atajo en la pestana Teclado.)

## Uso
1. Abre un **ensamblaje** o una **pieza multicuerpo** (guardada: la macro modifica el archivo) y ejecuta la macro.
2. Ventana unica (se coloca a la izquierda de SolidWorks): **AUTOMATICO** analiza todo. **MANUAL**: pulsa MANUAL,
   selecciona componentes (o, en una pieza, cuerpos) y pulsa CONTINUAR; solo se analizan las uniones entre las piezas elegidas.
   La seleccion es **acumulativa**: puedes volver a pulsar MANUAL, elegir otras piezas y CONTINUAR sin perder las anteriores
   (y se conserva lo que ya habias quitado a mano). **VACIAR SELECCION** empieza de cero. **LISTA - / LISTA +** cambian la altura de la ventana.
3. Aparece la **revision**: lista con una casilla por tarugo y cilindros temporales en el modelo (la vista pasa a alambre
   y se restaura al cerrar).
   - Marcar una casilla (o varias: clic, Mays+clic para un rango, Ctrl+A para todos) resalta ese tarugo (mas grande) y sus dos piezas en el 3D.
   - **FIJAR N**: escribe un numero (0 = automatico), marca algun tarugo de una union y pulsa FIJAR N: esa union lleva exactamente
     ese numero de tarugos (si no caben, se avisa en el motivo). Se recalcula conservando lo quitado a mano.
   - **QUITAR MARCADOS / RESTAURAR MARCADOS**, **MARCAR TODOS / DESMARCAR TODOS / MARCAR ELIMINADOS**, **QUITAR / RESTAURAR UNION COMPLETA**,
     **REACTIVAR TODOS** (vuelve a activar lo quitado). Un tarugo eliminado deja de dibujarse en el 3D. Las uniones dudosas empiezan desactivadas.
   - **Modo rapido** (casilla junto a ACEPTAR Y MECANIZAR): no repinta la pantalla ni actualiza el arbol mientras se taladra cada pieza; en la prueba (36 taladros) tardo 14 s frente a 57 s. Con ella marcada no se ve girar la pieza. La pantalla se restaura al terminar cada pieza. Por defecto desmarcada.
   - Al mecanizar aparece una ventana de **progreso** con barra de completado y boton **PARAR**. PARAR termina el taladro en curso, guarda la pieza con lo ya hecho y omite el resto; al volver a ejecutar continua (lo existente no se duplica). Hasta completar, los tarugos pendientes de una pieza se ven como `HECHO EN P1/P2` (solo una de las dos piezas taladrada). Si se cambia de documento activo mientras mecaniza, tambien se para.
   - El informe `Informe_Fase5.txt` incluye una linea `TIEMPOS` (busqueda de caras, creacion de taladros, calculo de volumen).
4. **ACEPTAR Y MECANIZAR** pide confirmacion, hace copia de seguridad (carpeta `Respaldo_Tarugos_<fecha>` junto al archivo),
   crea los taladros y guarda cada archivo una vez.
5. **DESHACER MECANIZADO** elimina solo las operaciones creadas por la macro (en ensamblaje y en pieza).

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

Numero de tarugos de una union de longitud L:
- L menor que el incremento de longitud (300): **1** tarugo centrado.
- L hasta la longitud de referencia: **proporcional**, N = redondeo(N inicial x L / Lref), entre 1 y N inicial
  (con Lref 1000 y N inicial 3: 1000 mm = 3, 500 mm = 2, 300 mm = 1).
- L mayor que la referencia: N = N inicial + IncN x ENTERO((L - Lref) / IncLong).
Separacion minima entre tarugos: 50 mm (6,25 x diametro). Si en una union
corta no caben N tarugos, se prueba con uno menos hasta 1 tarugo centrado (anotado en el motivo de la union).
Si un tarugo cae en un hueco de alguna de las piezas, se desplaza (manteniendo la distancia a los extremos) o, si no hay
sitio, se descarta en las dos piezas y se marca "SIN MATERIAL". Los ajustes se guardan en el registro de Windows
(`HKCU\Software\VB and VBA Program Settings\TarugosPro`).

## Resultados
- Operaciones nombradas `UNION_001_TARUGO_002_<pieza>_D8_P31-11`, con su posicion en el comentario (`TARUGO_META=...`).
- Informes `Informe_Fase5.txt` e `Informe_Tarugos.csv` (separador `;`) en la carpeta de respaldo.
- Volver a ejecutar no duplica: los taladros ya creados se reconocen y se omiten.
- En una pieza multicuerpo todos los taladros van en el mismo archivo y cada corte afecta solo al cuerpo de su pieza.

## Modificar o eliminar taladros
- Editar: en el arbol, clic derecho en la operacion `UNION_...` > Editar operacion.
- Eliminar todos los de la macro: DESHACER MECANIZADO. No se tocan operaciones que no lleven la marca de la macro.

## Piezas repetidas (solo ensamblajes)
Si una misma pieza aparece varias veces y las instancias necesitan taladros distintos, la macro pregunta:
SI = modificar el original (afecta a todas), NO = copia independiente (`<pieza>_T2.SLDPRT`) sustituida en este
ensamblaje, CANCELAR = omitir.

## Limitaciones conocidas
- **Uniones inclinadas (p. ej. ingletes a 45 grados): la macro no las taladra, a proposito.** No aparecen en la lista.
- Probada con tableros rectangulares a 90 grados, hasta unas decenas de piezas y SolidWorks 2024.
- Sin probar: ensamblajes cuyos componentes sean multicuerpo, contornos curvos, el mueble real con la version actual y SolidWorks 2025/2026.
- Uniones canto con canto quedan "dudosas".
- Piezas **virtuales** (guardadas dentro del ensamblaje, p. ej. por simetria) se omiten: guardalas antes como archivo externo.
- Solo hay un modo de analisis (rapido, con cajas envolventes).
- Mecanizar abre cada pieza con su ventana (SolidWorks necesita el documento activo para croquizar) y cierra las que abrio la macro.
- El clic en el 3D sobre un tarugo no es fiable: la lista es la forma de elegir.
- Deshacer solo borra: no restaura desde la copia de seguridad ni revierte la sustitucion por copias independientes.
- PDM: no integrado. Todo el acceso a archivos esta en `mGuardado` para anadirlo mas adelante.
- La ventana de revision no es modal; si SolidWorks se queda congelado, cerrar el proceso.
- Despues de ejecutar la macro, el editor VBA a veces no escribe el `.swp` en disco: para cambiar el codigo,
  reiniciar SolidWorks antes de abrir el editor y comprobar el guardado.

## Modulos
mConfig (parametros) - mVec, mPoligono (geometria) - mGeometria (piezas, cuerpos y caras) - mUniones (deteccion) -
mHuecos, mPosiciones (reparto) - mPreview, mEdicion, cEventos, frmRevision (revision) - mTaladros (cortes reales) -
mGuardado (acceso a archivos) - mDuplicados (marcas y duplicados) - mRecuperacion (deshacer) - mOptim (progreso) -
mInformes (CSV) - mMain (punto de entrada).
Revision de seguridad y limites de uso: `SEGURIDAD.md`.
Pruebas: `pruebas/` (generador del ensamblaje de pruebas por COM).
