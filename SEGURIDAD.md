# Tarugos Pro - revision de seguridad y robustez del codigo

Documento para la revision interna. Describe que hace y que NO hace la macro, que protecciones tiene y que limites conoce.
Version revisada: ver `git log` (ultimo commit). Probada en SolidWorks 2024.

## 1. Alcance: que toca la macro
- **Lee** la geometria del ensamblaje / pieza activos (API de SolidWorks).
- **Escribe**, solo tras confirmacion expresa del usuario (cuadro "Continuar? Si/No", por defecto NO):
  - cortes extruidos circulares (con su croquis) en los archivos de las piezas afectadas, agrupados en la carpeta `TARUGOS MACRO`;
  - una carpeta `Respaldo_Tarugos_<fecha>` junto al ensamblaje con una copia de cada archivo ANTES de modificarlo, mas `Informe_Fase5.txt` e `Informe_Tarugos.csv`;
  - `Tarugos_Uniones.txt` (diagnostico de uniones) junto al documento (o en `%TEMP%` si no tiene ruta);
  - los ajustes del usuario (9 numeros) en `HKCU\Software\VB and VBA Program Settings\TarugosPro`.
- **No hace**: acceso a red, ejecucion de programas (`Shell`, `CreateObject`, `WScript`), `SendKeys`, borrado de archivos (`Kill`, `RmDir`),
  lectura de otros datos del registro, ni cambia atributos de archivos. Comprobado con busqueda en todo el codigo.
- Llamadas a la API de Windows (solo en la ventana de revision): `GetWindowRect`, `GetDC`, `ReleaseDC`, `GetDeviceCaps`; son de **solo lectura**
  y sirven unicamente para colocar la ventana junto a la de SolidWorks.

## 2. Protecciones existentes
| Riesgo | Proteccion |
|---|---|
| Modificar sin querer | Nada se modifica hasta ACEPTAR Y MECANIZAR + confirmacion con el listado de archivos. La revision 3D solo dibuja cuerpos temporales. |
| Perder el original | Copia de seguridad de cada archivo antes de tocarlo. **Si la carpeta de respaldo no se puede crear, no se mecaniza nada** (mecanizado y deshacer). Si la copia de una pieza falla, esa pieza se omite. |
| Mecanizar otro documento | La ventana no es modal; antes de mecanizar se comprueba que el documento activo es el analizado (`AsmAnalizado`). |
| Archivos de solo lectura / sin extraer / virtuales | Se omiten y se anotan en el informe. |
| Duplicados | Cada operacion lleva la marca `TARUGO_META`; volver a ejecutar no duplica. |
| Corte incorrecto | Se verifica el volumen eliminado (aviso si se aparta de lo esperado) y que exista material. |
| Error inesperado | `On Error GoTo` en el mecanizado: se escribe el informe y se avisa; las piezas ya guardadas lo estan y hay copia. |
| Valores absurdos | `mConfig.ValidaParametros` limita todos los parametros (diametro 1-30, profundidades 1-200, N inicial 1-20 entero, holgura 0,001-5...). Solo se aceptan cifras y un separador decimal. Los ajustes del registro se validan al leerlos; si estan fuera de rango se usan los de fabrica. |
| N manual absurdo | Entero 0-20; si no caben, se reduce y se avisa en el motivo. |
| Parar el mecanizado | Ventana de progreso con PARAR (`frmProgreso`). Se comprueba una vez por taladro: nunca corta a mitad de un corte; guarda lo hecho y se puede reanudar sin duplicar. Si el documento activo cambia durante el mecanizado, se para solo. Si la ventana falla, el mecanizado se detiene (fallo seguro). |
| Analisis largo | Barra de progreso y cancelacion con Escape (`mOptim`); un analisis cancelado no modifica nada. |

## 3. Cambios de esta revision
- Validacion central de parametros (limites + registro).
- Sin carpeta de respaldo -> no se modifica nada (antes se intentaba seguir y fallaba la copia pieza a pieza).
- Eliminados del proyecto el formulario antiguo `frmPrincipal` y el modulo de ensamblajes de prueba `mPrueba` (codigo sin uso; el generador de pruebas queda en `pruebas/`).
- Recalculos de la ventana: lo quitado/restaurado a mano se conserva por clave (pieza + posicion) solo cuando se pide (MANUAL acumulativo, FIJAR N).

## 4. Limites conocidos / recomendaciones
- **`On Error Resume Next`** se usa en muchas consultas a la API (SolidWorks devuelve `Nothing`/vacio en lugar de error). Los resultados criticos
  (crear corte, guardar) se comprueban y se informan; el resto degrada a "no hecho" y nunca a "hecho".
- La macro **no firma** el proyecto VBA ni lo protege con contrasena. Para distribuir: firmar con certificado de la empresa y bloquear el proyecto VBA.
- Entre analisis y mecanizado el usuario podria cambiar el modelo (no se detecta salvo cambio de documento): repetir el analisis si se edita.
- Sin probar: SolidWorks 2025/2026, PDM (no integrado), ensamblajes con componentes multicuerpo.
- Registro: los ajustes son por usuario (HKCU); no se leen ni escriben otras claves.
- Se recomienda una revision de tres personas del modulo `mTaladros` (es el unico que escribe en los archivos de pieza) y probar siempre sobre una copia.
