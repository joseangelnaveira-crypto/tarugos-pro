# Crea el ensamblaje de PRUEBA "Mueble_Pruebas" en C:\Mueble_Pruebas (no toca C:\Mueble). Cotas en mm -> API en metros.
import os, sys, pythoncom, win32com.client as w
from win32com.client import VARIANT
def val(x): return x() if callable(x) else x
def byi(): return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
NONE = VARIANT(pythoncom.VT_DISPATCH, None)
M = 0.001
app = w.Dispatch(pythoncom.GetActiveObject("SldWorks.Application").QueryInterface(pythoncom.IID_IDispatch))
CARP = r"C:\Mueble_Pruebas"
os.makedirs(CARP, exist_ok=True)
tplP = app.GetUserPreferenceStringValue(8)
tplA = app.GetUserPreferenceStringValue(9)

def rect(sm, dx, dy):  sm.CreateCornerRectangle(0, 0, 0, dx*M, dy*M, 0)
def poligono(sm, pts):
    for i in range(len(pts)):
        a = pts[i]; b = pts[(i+1) % len(pts)]
        sm.CreateLine(a[0]*M, a[1]*M, 0, b[0]*M, b[1]*M, 0)

def crea_pieza(nombre, dz, perfil):
    ruta = os.path.join(CARP, nombre + ".SLDPRT")
    pm = app.NewDocument(tplP, 0, 0, 0)
    f = pm.FirstFeature; plano = None
    while f is not None:
        if val(f.GetTypeName2) == "RefPlane" and plano is None: plano = f
        f = f.GetNextFeature
    plano.Select2(False, 0)
    pm.SketchManager.InsertSketch(True)
    perfil(pm.SketchManager)
    pm.SketchManager.InsertSketch(True)
    f = pm.FirstFeature; sk = None
    while f is not None:
        if val(f.GetTypeName2) == "ProfileFeature": sk = f
        f = f.GetNextFeature
    sk.Select2(False, 0)
    ext = pm.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, dz*M, 0, False, False, False, False, 0, 0,
                                              False, False, False, False, True, True, True, 0, 0, False)
    ok = pm.Extension.SaveAs(ruta, 0, 1, NONE, byi(), byi())
    t = val(pm.GetTitle)
    print("pieza", nombre, "| extrusion:", ext is not None, "| guardada:", os.path.exists(ruta))
    app.CloseDoc(t)
    return ruta

# (nombre, dx, dy, dz, perfil, esquina minima X,Y,Z en el ensamblaje)
def r(dx, dy): return lambda sm: rect(sm, dx, dy)
def base(sm):
    rect(sm, 800, 400); sm.CreateCircleByRadius(409*M, 190*M, 0, 30*M)      # agujero de 60 mm bajo el divisor 1
def balda_muesca(sm):
    poligono(sm, [(0,0),(764,0),(764,380),(0,380),(0,250),(20,250),(20,150),(0,150)])   # muesca en el canto izquierdo
piezas = [
 ("P_Base",        800, 400, 18,  base,            (0,   0,   0)),
 ("P_Costado",      18, 400, 600, r(18,400),        (0,   0,   18)),
 ("P_Trasera",     800,  18, 618, r(800,18),        (0,   400, 0)),
 ("P_BaldaMuesca", 764, 380, 18,  balda_muesca,     (18,  0,   250)),
 ("P_Balda2",      764, 380, 18,  r(764,380),       (18,  20,  400)),
 ("P_Divisor1",     18, 360, 200, r(18,360),        (400, 20,  18)),
 ("P_Divisor2corto",18,  80, 200, r(18,80),         (600, 100, 18)),
 ("P_ChapaFina",     2, 300, 150, r(2,300),         (700, 50,  18)),
 ("P_PanelA",       18, 150, 200, r(18,150),        (200, 100, 18)),
 ("P_PanelB",       18, 150, 200, r(18,150),        (200, 250, 18)),
 ("P_Interferencia",100, 100, 18, r(100,100),       (100, 100, 405)),
]
rutas = {}
for nom, dx, dy, dz, per, mn in piezas:
    rutas[nom] = crea_pieza(nom, dz, per)

# ensamblaje
for nom in rutas:
    app.OpenDoc6(rutas[nom], 1, 1, "", byi(), byi())
asm = app.NewDocument(tplA, 0, 0, 0)
n = 0
def pon(nom, dx, dy, dz, mn):
    global n
    cx, cy, cz = (mn[0]+dx/2)*M, (mn[1]+dy/2)*M, (mn[2]+dz/2)*M
    c = asm.AddComponent5(rutas[nom], 0, "", False, "", cx, cy, cz)
    print("componente", nom, "->", c is not None)
    n += c is not None
for nom, dx, dy, dz, per, mn in piezas:
    pon(nom, dx, dy, dz, mn)
    if nom == "P_Costado":
        pon(nom, dx, dy, dz, (782, 0, 18))      # 2.a instancia del MISMO archivo (costado derecho)
ens = os.path.join(CARP, "Mueble_Pruebas.SLDASM")
asm.Extension.SaveAs(ens, 0, 1, NONE, byi(), byi())
asm.ViewZoomtofit2
print("componentes:", n, "| ensamblaje guardado:", os.path.exists(ens))
