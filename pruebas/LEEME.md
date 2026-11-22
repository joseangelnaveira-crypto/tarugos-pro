# Ensamblaje de prueba (casos dificiles)

`crea_pruebas.py` crea en `C:\Mueble_Pruebas` un ensamblaje `Mueble_Pruebas.SLDASM` con 12 componentes
(11 archivos) para probar la macro sin tocar tus disenos. Necesita Python con pywin32 y SolidWorks abierto:

    python pruebas\crea_pruebas.py

Casos incluidos (cotas en mm; en la vista "Frontal" de SolidWorks el eje vertical es Y):

| Pieza | Caso que prueba | Resultado esperado |
|---|---|---|
| P_Base (agujero de 60 mm) | contacto con agujero interior bajo P_Divisor1 | un tarugo se desplaza (>= 50 mm del borde del agujero) |
| P_BaldaMuesca | muesca de 100 mm en un canto | el contacto se parte en dos uniones cortas |
| P_Divisor2corto (80 mm) | union corta | 1 tarugo centrado |
| P_ChapaFina (2 mm) | canto muy estrecho | union DUDOSA, sin taladros |
| P_PanelA / P_PanelB | canto con canto | union DUDOSA, sin taladros |
| P_Costado (2 instancias) | pieza repetida con taladros distintos | pregunta de pieza compartida |
| P_Interferencia | pieza solapada con otra | sin union (limitacion conocida) |

Resultado de la primera ejecucion (22/11/2026): 90 de 102 taladros, 0 errores, 90 comprobaciones de volumen correctas.
