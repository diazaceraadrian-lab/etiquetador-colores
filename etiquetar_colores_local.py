# -*- coding: utf-8 -*-
r"""
ETIQUETADOR DE COLORES - version local, SIN IA y SIN CLAVES DE API
===================================================================

Por cada producto del listado: descarga su foto, separa el producto del
fondo blanco de estudio, calcula su color dominante y lo traduce al
nombre comercial en ingles mas parecido.

Todo se ejecuta en tu ordenador. No usa OpenAI ni ningun otro servicio de
pago: coste cero y sin credenciales que se puedan filtrar.

COMO SE USA (lo mas sencillo primero)
-------------------------------------
1) Arrastra el listado encima de "ETIQUETAR COLORES.bat".
   Puedes arrastrar varios a la vez.
2) O haz doble clic en "ETIQUETAR COLORES.bat" y elige el fichero.
3) O desde una ventana de comandos:
       python etiquetar_colores_local.py "mis_productos.csv"
       python etiquetar_colores_local.py listado1.csv listado2.xlsx

QUE FORMATOS ACEPTA
-------------------
.csv (separado por ; , tabulador o |) y .xlsx / .xls. Encuentra las
columnas por su CONTENIDO, no por el nombre, asi que da igual como vengan
las cabeceras y da igual que la columna de las fotos venga sin nombre.
Tambien arregla las URLs que llegan con las barras escapadas al estilo
JSON, que si no fallarian todas (le pasaba a todas las filas de
un listado real).

QUE GENERA
----------
Un .xlsx junto al fichero de entrada, llamado "<lo mismo> CON COLOR.xlsx",
con tus columnas originales mas estas seis:
    Color      nombre comercial en ingles (Black, Navy blue, Camel...)
    Familia    grupo de color (black, blue, brown...), util para filtrar
    Confianza  0 a 1: que parte del producto es de ese color
    Revisar    SI / NO: si conviene que lo mire una persona
    Motivo     por que se marca para revisar
    RGB        el color medido, en formato #RRGGBB

Las fotos se guardan en la subcarpeta "cache_imagenes": la segunda vez que
proceses el mismo listado tarda segundos en lugar de minutos. Si lo paras
a media, al relanzarlo continua donde lo dejo. Puedes borrar
"cache_imagenes" cuando quieras para liberar espacio.

Ritmo real medido: 1042 productos en 3.5 minutos descargando todo; unos
40 segundos si las fotos ya estan en la cache.

PRECISION MEDIDA (medida, no estimada)
--------------------------------------
Validado contra 237 fotos de dos proveedores, con cuatro conjuntos de
control: el color que el proveedor escribe dentro de la referencia (136 +
21 casos) y familias de color juzgadas mirando las fotos una a una (40 +
40 casos).
    familia de color exacta ........... 81%
    exacta o familia vecina ........... 97%
    fallos claros .....................  3%   (7 de 237)

De esos 7 fallos: 3 son objetos DORADOS METALIZADOS, 2 son etiquetas mal
puestas por el proveedor (una dice "sage" y la prenda es gris calido;
otra dice "marrone" y la camisa es de cuadros grises), 1 es un boligrafo
y 1 es un zapato nude etiquetado como rosa.

"Familia vecina" significa confundir tonos que a la vista se parecen de
verdad: beige con marron, gris con negro, rosa con morado. Para un
catalogo suele ser aceptable; para el 81% de acierto exacto, el nombre
es el que pondrias tu.

LIMITES CONOCIDOS Y COMPROBADOS
-------------------------------
1. METALIZADOS (oro, plata): salen como beige, crema o gris claro. En una
   foto, el oro y el crema son literalmente el mismo color; solo los
   distingue el brillo. Se probo a detectarlos midiendo el contraste
   especular y NO funciona: los dorados dieron valores de 1.8 a 64.1 y
   los tejidos mates de 25.7 a 77.1, completamente solapados. No hay
   aviso posible: hay que mirarlos a mano. Truco: busca en la columna de
   referencia las palabras ORO, GOLD, SILVER, METALLIC.
2. VAQUEROS MUY LAVADOS: promedian a un gris casi neutro y pueden salir
   "Charcoal" o "Dark grey" en vez de azul. Tampoco es separable: los
   grises reales del catalogo llegan al mismo valor.
3. ESTAMPADOS de tonos parecidos (floral marron sobre beige): devuelve el
   color de fondo, que normalmente es el correcto para un catalogo, pero
   el aviso de multicolor no siempre salta.
4. Si en la foto hay DOS prendas (un abrigo sobre unos vaqueros), mide el
   color dominante y no sabe cual de las dos es el producto.
5. Productos que no son ropa (boligrafos, perfumes) van peor: la forma no
   se parece a lo que esto espera.

Revisar=SI no significa que el color este mal: significa que hay algo en
la foto que reduce la certeza.

QUE HACER SI FALLA
------------------
El programa explica en castellano que ha pasado. Los casos tipicos:
 - "No existe el fichero": la ruta esta mal o el fichero se movio.
 - "No encuentro ninguna columna con enlaces a las fotos": el listado no
   trae URLs de imagen; pidelo con la columna de la foto incluida.
 - "el fichero esta abierto": cierra el Excel de salida y relanza.
 - Si sale "ERROR_DESCARGA" en algunas filas, esas URLs dan 404 en el
   servidor del proveedor. No es cosa del programa.
"""

import io
import os
import re
import sys

import numpy as np
import requests
from PIL import Image
from scipy import ndimage

# ---------------------------------------------------------------
# PALETA: (nombre comercial en ingles, RGB de referencia, familia)
# Un mismo nombre puede tener VARIOS anclajes RGB, para cubrir el mismo
# color en version viva y en version apagada. Sin los anclajes apagados,
# los marrones y verdes desaturados caian en el gris mas cercano.
# ---------------------------------------------------------------
PALETA = [
    ("Black",       (25, 25, 25),    "black"),
    ("Charcoal",    (62, 62, 64),    "grey"),
    ("Dark grey",   (98, 98, 101),   "grey"),
    ("Grey",        (142, 142, 145), "grey"),
    ("Light grey",  (196, 196, 199), "grey"),
    ("Off white",   (234, 231, 224), "white"),
    ("White",       (251, 251, 251), "white"),
    ("Cream",       (241, 233, 210), "white"),
    ("Beige",       (222, 205, 178), "beige"),
    ("Beige",       (205, 193, 178), "beige"),   # nude / arena apagado
    ("Sand",        (205, 185, 155), "beige"),
    ("Sand",        (212, 180, 128), "beige"),   # paja / rafia natural
    ("Taupe",       (166, 152, 136), "beige"),
    ("Taupe",       (128, 116, 105), "beige"),   # taupe oscuro apagado
    ("Camel",       (186, 146, 100), "brown"),
    ("Tan",         (158, 119, 80),  "brown"),
    ("Brown",       (110, 76, 50),   "brown"),
    ("Brown",       (112, 94, 82),   "brown"),   # marron apagado
    ("Chocolate",   (84, 56, 40),    "brown"),
    ("Dark brown",  (62, 44, 33),    "brown"),
    ("Dark brown",  (80, 67, 62),    "brown"),   # marron oscuro apagado
    ("Dark brown",  (42, 28, 22),    "brown"),   # casi negro pero calido
    ("Dark brown",  (52, 45, 40),    "brown"),   # marron oscuro casi neutro
    ("Dark brown",  (55, 47, 44),    "brown"),   # marron grisaceo oscuro
    ("Khaki",       (150, 140, 100), "beige"),
    ("Olive green", (110, 110, 60),  "green"),
    ("Olive green", (92, 90, 64),    "green"),   # militar medio apagado
    ("Olive green", (50, 50, 38),    "green"),   # militar muy oscuro
    ("Dark green",  (34, 68, 50),    "green"),
    ("Green",       (60, 130, 75),   "green"),
    ("Light green", (140, 190, 140), "green"),
    ("Mint green",  (170, 220, 195), "green"),
    ("Teal",        (38, 120, 120),  "green"),
    ("Navy blue",   (30, 42, 78),    "blue"),
    ("Navy blue",   (44, 52, 72),    "blue"),    # azul marino apagado
    ("Denim blue",  (85, 110, 150),  "blue"),
    ("Blue",        (45, 85, 172),   "blue"),
    ("Royal blue",  (34, 68, 192),   "blue"),
    ("Sky blue",    (110, 175, 225), "blue"),
    ("Light blue",  (168, 200, 228), "blue"),
    ("Purple",      (100, 60, 140),  "purple"),
    ("Lilac",       (185, 165, 215), "purple"),
    ("Burgundy",    (95, 30, 45),    "red"),
    ("Burgundy",    (96, 45, 41),    "red"),     # burdeos apagado
    ("Dark red",    (132, 26, 30),   "red"),
    ("Red",         (202, 35, 40),   "red"),
    ("Coral",       (240, 110, 95),  "red"),
    ("Fuchsia",     (215, 45, 130),  "pink"),
    ("Pink",        (235, 140, 170), "pink"),
    ("Light pink",  (245, 205, 215), "pink"),
    ("Pink",        (216, 150, 148), "pink"),   # rosa salmon / tea rose
    #  ^ medido: con b* mas alto (mas naranja) este anclaje robaba los camel
    ("Dusty pink",  (204, 170, 165), "pink"),   # rosa empolvado
    ("Orange",      (235, 120, 35),  "orange"),
    ("Mustard",     (205, 165, 50),  "yellow"),
    ("Yellow",      (240, 215, 60),  "yellow"),
]

# Rampa de neutros: cuando el color no tiene tinte, solo decide la luz
# El corte Negro/Charcoal en L=26 sale de barrer 20/22/24/26/28 contra los
# TRES conjuntos de control (67 + 40 + 21 fotos): 26 es el mejor en los tres
# a la vez, 83.6% -> 87.5% de acierto exacto, y sin anadir ningun fallo.
# Un negro con brillo (piel lisa, charol) sube de luminosidad y con el corte
# en 20 se iba a "Charcoal".
RAMPA_NEUTROS = [
    (26, "Black"), (38, "Charcoal"), (52, "Dark grey"),
    (68, "Grey"), (84, "Light grey"), (93, "Off white"), (101, "White"),
]

FAMILIA_DE = {n: f for n, _, f in PALETA}
FAMILIA_DE.update({"Black": "black", "Charcoal": "grey", "Dark grey": "grey",
                   "Grey": "grey", "Light grey": "grey",
                   "Off white": "white", "White": "white"})

# --- Umbrales ---
# medido: en los 51 negros/grises/blancos correctos del catalogo, el eje a*
# (calido-frio) nunca paso de 2.1; en los marrones que se confundian con
# gris, siempre fue >= 5.0. De ahi el corte en 3.0.
CROMA_NEUTRO_SEGURO = 4.0
CROMA_NEUTRO_MAX = 9.0
A_NEUTRO_MAX = 3.0
B_NEUTRO_MAX = 4.0

DE_MISMO_COLOR = 17.0        # por debajo: mismo color con distinta luz
DE_TONO_MULTICOLOR = 13.0    # dif. minima de TONO para avisar de multicolor
PESO_MIN_MULTICOLOR = 0.26

CONFIANZA_MIN = 0.45         # por debajo: avisar de color poco dominante
MAX_AGUJERO = 0.015          # agujeros mayores no se rellenan
ENCERRADO_MIN = 0.015        # zona clara encerrada relevante (% de la foto)
# medido: el hueco encerrado se descarta como fondo solo si es PEQUENO en
# relacion al cuerpo del producto (ratio < 1). Con ratio alto no se puede
# distinguir del blanco propio del producto (zapatillas blancas: 5.4-7.6;
# collar de oro: 8.4) y descartarlo estropeaba los productos blancos.
RATIO_ENCERRADO_MAX = 1.0
AVISO_ENCERRADO = 0.03       # a partir de aqui, avisar de producto hueco

_M = np.array([[0.4124564, 0.3575761, 0.1804375],
               [0.2126729, 0.7151522, 0.0721750],
               [0.0193339, 0.1191920, 0.9503041]])
_BLANCO_D65 = np.array([0.95047, 1.00000, 1.08883])


def rgb_a_lab(rgb):
    """Pasa color RGB (0-255) a CIE Lab, donde la distancia entre dos
    colores se corresponde con la diferencia que ve el ojo."""
    r = np.asarray(rgb, dtype=np.float64) / 255.0
    lineal = np.where(r <= 0.04045, r / 12.92, ((r + 0.055) / 1.055) ** 2.4)
    xyz = lineal @ _M.T / _BLANCO_D65
    e, k = 216.0 / 24389.0, 24389.0 / 27.0
    f = np.where(xyz > e, np.cbrt(xyz), (k * xyz + 16.0) / 116.0)
    return np.stack([116.0 * f[..., 1] - 16.0,
                     500.0 * (f[..., 0] - f[..., 1]),
                     200.0 * (f[..., 1] - f[..., 2])], axis=-1)


def lab_a_rgb(lab):
    lab = np.asarray(lab, dtype=np.float64)
    fy = (lab[..., 0] + 16.0) / 116.0
    f = np.stack([fy + lab[..., 1] / 500.0, fy, fy - lab[..., 2] / 200.0],
                 axis=-1)
    e, k = 216.0 / 24389.0, 24389.0 / 27.0
    xyz = np.where(f ** 3 > e, f ** 3, (116.0 * f - 16.0) / k) * _BLANCO_D65
    lineal = xyz @ np.linalg.inv(_M).T
    srgb = np.where(lineal <= 0.0031308, 12.92 * lineal,
                    1.055 * np.maximum(lineal, 0) ** (1 / 2.4) - 0.055)
    return np.clip(srgb * 255.0, 0, 255).astype(np.uint8)


PALETA_LAB = rgb_a_lab(np.array([c for _, c, _ in PALETA], dtype=np.float64))
PALETA_NOMBRES = [n for n, _, _ in PALETA]


def delta_e(lab1, lab2):
    """Diferencia visual total entre dos colores (CIE76)."""
    return np.linalg.norm(np.asarray(lab1) - np.asarray(lab2), axis=-1)


def delta_tono(lab1, lab2):
    """Diferencia solo de tono, ignorando la luz. Evita confundir
    'negro con brillo de charol' con 'blanco y negro'."""
    d = np.asarray(lab1)[..., 1:] - np.asarray(lab2)[..., 1:]
    return np.linalg.norm(d, axis=-1)


def nombre_color(lab):
    """Traduce un color Lab al nombre comercial mas parecido."""
    L, a, b = float(lab[0]), float(lab[1]), float(lab[2])
    croma = (a * a + b * b) ** 0.5
    es_neutro = (croma < CROMA_NEUTRO_SEGURO
                 or (croma < CROMA_NEUTRO_MAX
                     and abs(a) < A_NEUTRO_MAX and abs(b) < B_NEUTRO_MAX))
    if es_neutro:
        for limite, nombre in RAMPA_NEUTROS:
            if L < limite:
                return nombre
        return "White"
    return PALETA_NOMBRES[int(np.argmin(delta_e(PALETA_LAB,
                                                np.array([L, a, b]))))]


def familia_de(nombre):
    return FAMILIA_DE.get(nombre, "desconocida")


def descargar_imagen(url, timeout=20, intentos=2, lado=420):
    """Descarga, aplana transparencias sobre blanco y reduce a 420 px."""
    ultimo = None
    for _ in range(intentos):
        try:
            r = requests.get(url, timeout=timeout,
                             headers={"User-Agent": "Mozilla/5.0"})
            r.raise_for_status()
            img = Image.open(io.BytesIO(r.content))
            if img.mode in ("RGBA", "LA", "P"):
                img = img.convert("RGBA")
                fondo = Image.new("RGB", img.size, (255, 255, 255))
                fondo.paste(img, mask=img.split()[-1])
                img = fondo
            else:
                img = img.convert("RGB")
            img.thumbnail((lado, lado), Image.LANCZOS)
            return img
        except Exception as e:
            ultimo = e
    raise ultimo


def mascara_prenda(img, tolerancia=14):
    """
    Devuelve (mascara, info). mascara True = pixel del producto.

    El fondo es la zona clara y sin color conectada con el borde de la foto.
    Ademas, una zona clara GRANDE que el producto deja encerrada (el hueco
    de una cadena) tambien es fondo, pero solo se descarta si es pequena en
    relacion al cuerpo del producto: si no, no se distingue del blanco
    propio de un producto blanco.
    """
    a = np.asarray(img, dtype=np.int16)
    h, w = a.shape[:2]
    total = float(h * w)

    # Color del fondo estimado con las 4 esquinas (no siempre es blanco puro)
    k = max(3, min(h, w) // 40)
    esquinas = np.concatenate([
        a[:k, :k].reshape(-1, 3), a[:k, -k:].reshape(-1, 3),
        a[-k:, :k].reshape(-1, 3), a[-k:, -k:].reshape(-1, 3)])
    fondo_rgb = np.median(esquinas, axis=0)

    dist = np.abs(a - fondo_rgb).max(axis=2)
    saturacion = a.max(axis=2) - a.min(axis=2)
    candidato = (dist <= tolerancia) & (saturacion <= 18)

    info = {"encerrado": 0.0, "hueco_descartado": False}
    etiquetas, n = ndimage.label(candidato)
    if n == 0:
        bruta = np.ones((h, w), dtype=bool)
    else:
        borde = np.concatenate([etiquetas[0, :], etiquetas[-1, :],
                                etiquetas[:, 0], etiquetas[:, -1]])
        ids = set(int(x) for x in np.unique(borde) if x != 0)
        tam = np.bincount(etiquetas.ravel())

        grandes = [i for i in range(1, n + 1)
                   if i not in ids and tam[i] > ENCERRADO_MIN * total]
        area_encerrada = float(sum(tam[i] for i in grandes))
        cuerpo = float((~candidato).sum())     # parte no clara = producto
        ratio = area_encerrada / cuerpo if cuerpo else 999.0
        info["encerrado"] = area_encerrada / total

        if grandes and ratio < RATIO_ENCERRADO_MAX:
            ids_fondo = list(ids) + grandes
            info["hueco_descartado"] = True
        else:
            ids_fondo = list(ids)
        bruta = (~np.isin(etiquetas, ids_fondo) if ids_fondo
                 else np.ones((h, w), bool))

    # Rellenar solo agujeros pequenos (poros, costuras, reflejos puntuales)
    lleno = ndimage.binary_fill_holes(bruta)
    huecos = lleno & ~bruta
    if huecos.any():
        eh, _ = ndimage.label(huecos)
        tam_h = np.bincount(eh.ravel())
        pequenos = np.where(tam_h <= MAX_AGUJERO * total)[0]
        pequenos = pequenos[pequenos != 0]
        if len(pequenos):
            bruta = bruta | np.isin(eh, pequenos)

    area_bruta = bruta.mean()

    # Erosion adaptativa: quita el borde de mezcla producto/fondo sin
    # comerse piezas finas como una cadena o el tacon de un zapato
    prenda = bruta
    for tam_k in (5, 3):
        prueba = ndimage.binary_erosion(bruta, np.ones((tam_k, tam_k), bool))
        if prueba.mean() > 0.35 * area_bruta and prueba.mean() > 0.004:
            prenda = prueba
            break

    frac = prenda.mean()
    if frac < 0.003 or frac > 0.97:
        prenda = np.zeros((h, w), dtype=bool)
        prenda[int(h * 0.30):int(h * 0.70), int(w * 0.35):int(w * 0.65)] = True
        info["recorte_central"] = True
    return prenda, info


def mascara_piel(lab):
    """Rango tipico de piel humana. Se solapa con camel/beige, asi que solo
    se aplica cuando la piel es minoria (fotos con modelo)."""
    L, a, b = lab[:, 0], lab[:, 1], lab[:, 2]
    return ((L > 25) & (L < 88) & (a > 5) & (a < 26)
            & (b > 9) & (b < 34) & (b > a * 0.75))


def _kmeans(X, k, iteraciones=18, semilla=7):
    """k-means minimo en numpy (evita depender de scikit-learn)."""
    rng = np.random.default_rng(semilla)
    centros = [X[rng.integers(len(X))]]
    for _ in range(k - 1):
        d = np.min(np.stack([np.linalg.norm(X - c, axis=1) for c in centros]),
                   axis=0)
        prob = d ** 2
        s = prob.sum()
        idx = rng.integers(len(X)) if s <= 0 else rng.choice(len(X), p=prob / s)
        centros.append(X[idx])
    C = np.array(centros, dtype=np.float64)
    asign = np.full(len(X), -1, dtype=np.int64)
    for _ in range(iteraciones):
        d = np.linalg.norm(X[:, None, :] - C[None, :, :], axis=2)
        nueva = np.argmin(d, axis=1)
        if np.array_equal(nueva, asign):
            break
        asign = nueva
        for j in range(k):
            sel = asign == j
            if sel.any():
                C[j] = X[sel].mean(axis=0)
    return C, np.array([(asign == j).mean() for j in range(k)])


def detectar_color(img, k=4, max_pixeles=12000):
    """Devuelve un dict: color, familia, confianza, rgb, avisos."""
    mask, info = mascara_prenda(img)
    cobertura = float(mask.mean())
    pix = np.asarray(img, dtype=np.float64)[mask]
    base = {"color": "ERROR_SIN_PRENDA", "familia": None, "confianza": 0.0,
            "rgb": None, "cobertura": round(cobertura, 4), "piel": 0.0,
            "multicolor": False, "frontera_gris": False,
            "n_grupos": 0, "revisar": True,
            "motivo": "no se consigue aislar el producto del fondo"}
    if len(pix) < 40:
        return base

    if len(pix) > max_pixeles:
        rng = np.random.default_rng(7)
        pix = pix[rng.choice(len(pix), max_pixeles, replace=False)]

    lab = rgb_a_lab(pix)

    # NO se descartan los pixeles color piel. Medido sobre el catalogo:
    # cambiaba el color de 43 productos y el balance era negativo (una
    # cartera beige pasaba a "Dark green", un bolso marron a "White"),
    # porque el rango de la piel se solapa con camel/beige. Solo se mide
    # para dejarlo como dato informativo.
    frac_piel = float(mascara_piel(lab).mean())

    # Fuera brillos especulares y sombras profundas: falsean el tono
    L = lab[:, 0]
    filtro = (L > np.percentile(L, 4)) & (L < np.percentile(L, 96))
    if filtro.sum() > 200:
        lab = lab[filtro]

    kk = max(1, min(k, len(np.unique(lab.round(0), axis=0))))
    C, pesos = _kmeans(lab, kk)
    orden = np.argsort(-pesos)
    C, pesos = C[orden], pesos[orden]

    # Fundir los grupos que son el mismo color con distinta iluminacion
    juntos = [0] + [j for j in range(1, len(C))
                    if delta_e(C[j], C[0]) < DE_MISMO_COLOR]
    peso_principal = float(pesos[juntos].sum())
    color_lab = np.average(C[juntos], axis=0, weights=pesos[juntos])

    # Multicolor es un AVISO, no un color: el dominante se devuelve siempre
    multicolor = any(
        pesos[j] >= PESO_MIN_MULTICOLOR
        and float(delta_tono(C[j], color_lab)) >= DE_TONO_MULTICOLOR
        for j in range(len(C)) if j not in juntos)

    nombre = nombre_color(color_lab)

    motivos = []
    if multicolor:
        motivos.append("posible estampado o bicolor")
    if cobertura < 0.02:
        motivos.append("producto muy pequeno o fino (joyeria)")
    if info.get("encerrado", 0) > AVISO_ENCERRADO and not info["hueco_descartado"]:
        motivos.append("producto hueco: puede haber colado fondo blanco")
    if info.get("recorte_central"):
        motivos.append("no se aislo el producto, se uso el centro de la foto")
    if peso_principal < CONFIANZA_MIN:
        motivos.append("color poco dominante")
    a0, b0 = float(color_lab[1]), float(color_lab[2])
    croma0 = (a0 * a0 + b0 * b0) ** 0.5
    frontera = CROMA_NEUTRO_SEGURO <= croma0 < CROMA_NEUTRO_MAX

    base.update(color=nombre, familia=familia_de(nombre),
                confianza=round(peso_principal, 3),
                rgb=tuple(int(x) for x in lab_a_rgb(color_lab)),
                piel=round(frac_piel, 3), multicolor=multicolor,
                frontera_gris=frontera, n_grupos=int(len(C)),
                revisar=bool(motivos), motivo="; ".join(motivos))
    return base


# ===================================================================
#  LECTURA DEL LISTADO
#  Las columnas se detectan por CONTENIDO, no por su nombre, porque los
#  listados llegan con cabeceras distintas cada vez (y a veces la columna
#  de la URL viene sin nombre).
# ===================================================================
import argparse
import hashlib
import time
import traceback
from concurrent.futures import ThreadPoolExecutor

import pandas as pd

def _carpeta_del_programa():
    """
    Carpeta donde vive el programa.

    Si esto se ha empaquetado en un .exe con PyInstaller, __file__ apunta a
    un directorio temporal que Windows borra al salir; hay que usar la ruta
    del propio .exe. Si no se tiene esto en cuenta, la cache de fotos se
    pierde en cada ejecucion y se vuelve a descargar todo.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# Carpetas que se sincronizan con la nube. Si el programa vive dentro de
# una de ellas, la cache NO debe ir a su lado: serian miles de fotos
# subiendo y bajando para todo el equipo.
CARPETAS_QUE_SINCRONIZAN = ("google drive", "googledrive", "mi unidad",
                            "my drive", "onedrive", "dropbox", "sharepoint",
                            "icloud", "pcloud", "mega sync", "megasync")


def _esta_en_carpeta_sincronizada(ruta):
    baja = str(ruta).lower().replace("\\", "/")
    return any(("/" + p) in baja or baja.startswith(p)
               for p in CARPETAS_QUE_SINCRONIZAN)


def _cache_local_del_usuario():
    base = (os.environ.get("LOCALAPPDATA")
            or os.environ.get("TEMP")
            or os.path.expanduser("~"))
    ruta = os.path.join(base, "etiquetador-colores", "cache_imagenes")
    os.makedirs(ruta, exist_ok=True)
    return ruta


def _elegir_carpeta_cache():
    """
    Decide donde guardar las fotos descargadas. Tres casos:

    1. El programa esta en una carpeta que se sincroniza con la nube
       (Google Drive, OneDrive, Dropbox...): la cache va a la carpeta local
       del usuario. Si fuera al lado del programa, cada foto descargada se
       subiria a la nube y se bajaria en el ordenador de cada compañero:
       miles de ficheros y gigas de trafico para nada.
    2. El programa esta en una carpeta de red normal donde se puede
       escribir: la cache va a su lado y se COMPARTE. Cuando un compañero
       ya ha procesado un listado, al siguiente le va en segundos.
    3. La carpeta del programa es de solo lectura: cache local del usuario.
    """
    carpeta = _carpeta_del_programa()
    if _esta_en_carpeta_sincronizada(carpeta):
        return _cache_local_del_usuario()

    candidata = os.path.join(carpeta, "cache_imagenes")
    try:
        os.makedirs(candidata, exist_ok=True)
        prueba = os.path.join(candidata, ".escritura_ok")
        with open(prueba, "w") as fh:
            fh.write("ok")
        os.remove(prueba)
        return candidata
    except Exception:
        return _cache_local_del_usuario()


CARPETA_SCRIPT = _carpeta_del_programa()
CACHE = _elegir_carpeta_cache()

NOMBRES_REF = ("REFERENCE", "REFERENCIA", "SKU", "REF", "CODIGO", "CODE",
               "ARTICULO", "ARTICLE", "MODELO", "MODEL")
NOMBRES_SUPPLIER = ("SUPPLIER_ID", "SUPPLIERID", "SUPPLIER", "PROVEEDOR",
                    "ID_PROVEEDOR", "PROVEEDOR_ID")
COLUMNAS_NUEVAS = ["Color", "Familia", "Confianza", "Revisar", "Motivo", "RGB"]

# ---- Sentencias SQL que se generan en el Excel -------------------------
# Son opcionales: solo se generan si se lanza con  --sql
# Los nombres reales de tu base de datos NO van escritos aqui: se leen de
# variables de entorno. Configuralas una sola vez en una terminal (y abre
# una terminal nueva despues), solo las que sean distintas del valor por
# defecto:
#   setx ETIQUETADOR_TABLA "nombre_de_tu_tabla"           (productos)
#   setx ETIQUETADOR_COL_COLOR "columna_del_color"        (color)
#   setx ETIQUETADOR_COL_PROVEEDOR "columna_proveedor"    (proveedor_id)
#   setx ETIQUETADOR_COL_REFERENCIA "columna_referencia"  (referencia)
# Tambien se puede pasar otra tabla al lanzarlo con  --tabla otra_tabla
TABLA_SQL = os.environ.get("ETIQUETADOR_TABLA", "productos")
COLUMNA_SQL = os.environ.get("ETIQUETADOR_COL_COLOR", "color")
COLUMNA_SQL_PROVEEDOR = os.environ.get("ETIQUETADOR_COL_PROVEEDOR", "proveedor_id")
COLUMNA_SQL_REFERENCIA = os.environ.get("ETIQUETADOR_COL_REFERENCIA", "referencia")


def _sql_cadena(v):
    """Escapa un texto para meterlo entre comillas simples en SQL.
    Una comilla simple se duplica: es como se escapa en SQL estandar."""
    return str(v).replace("'", "''")


def _sql_supplier(v):
    """El proveedor va sin comillas si es un numero, como en
    'proveedor_id = 7'. Si viniera como texto, se entrecomilla."""
    s = str(v).strip()
    return s if re.fullmatch(r"-?\d+", s) else "'%s'" % _sql_cadena(s)


def sentencia_update(color, supplier, referencia, tabla=None):
    """Devuelve el UPDATE de una fila, listo para pegar en el gestor."""
    return ("update %s set %s = '%s' where %s = %s and %s = '%s';"
            % (tabla or TABLA_SQL, COLUMNA_SQL, _sql_cadena(color),
               COLUMNA_SQL_PROVEEDOR, _sql_supplier(supplier),
               COLUMNA_SQL_REFERENCIA, _sql_cadena(referencia)))


def limpiar_url(u):
    """
    Deja la URL utilizable.

    Los listados llegan a veces con las barras escapadas al estilo JSON
    ('https:\\/\\/www...') y asi la libreria de descarga las rechaza.
    Comprobado: en un listado real le pasaba a todas las filas.
    """
    if u is None:
        return ""
    u = str(u).strip().strip('"').strip("'")
    u = u.replace("\\/", "/").replace("\\\\", "/")
    u = re.sub(r"\s+", "", u)
    for esquema in ("http", "https"):
        if u.startswith(esquema + ":/") and not u.startswith(esquema + "://"):
            u = u.replace(esquema + ":/", esquema + "://", 1)
    return u


def _parece_url(serie):
    s = serie.astype(str).map(limpiar_url)
    return s.str.startswith(("http://", "https://")).mean()


def leer_listado(ruta):
    """Devuelve (tabla, columna_referencia, columna_url)."""
    if not os.path.exists(ruta):
        raise ValueError("No existe el fichero:\n  %s" % ruta)
    if os.path.isdir(ruta):
        raise ValueError("Eso es una carpeta, no un fichero:\n  %s" % ruta)
    if os.path.getsize(ruta) == 0:
        raise ValueError("El fichero esta vacio:\n  %s" % ruta)

    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        d = pd.read_excel(ruta, dtype=str)
    else:
        mejor = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            for sep in (";", ",", "\t", "|"):
                try:
                    p = pd.read_csv(ruta, sep=sep, encoding=enc, dtype=str,
                                    engine="python", on_bad_lines="skip")
                except Exception:
                    continue
                if p.shape[1] < 2:
                    continue
                # Gana la combinacion con mas columnas y con alguna que
                # tenga pinta de contener enlaces
                puntos = p.shape[1] + 10 * max(
                    (_parece_url(p[c]) for c in p.columns), default=0)
                if mejor is None or puntos > mejor[0]:
                    mejor = (puntos, p)
            if mejor and mejor[0] >= 12:
                break
        if mejor is None:
            raise ValueError(
                "No consigo interpretar el fichero como tabla.\n"
                "  Comprueba que sea un .csv o un .xlsx de verdad.")
        d = mejor[1]

    d.columns = [str(c).strip() for c in d.columns]
    d = d.reset_index(drop=True)

    col_url, mejor_frac = None, 0.0
    for c in d.columns:
        frac = _parece_url(d[c])
        if frac > mejor_frac:
            mejor_frac, col_url = frac, c
    if col_url is None or mejor_frac < 0.5:
        raise ValueError(
            "No encuentro ninguna columna con enlaces a las fotos.\n"
            "  Columnas del fichero: %s" % list(d.columns))

    col_ref = None
    for c in d.columns:
        if c.upper() in NOMBRES_REF:
            col_ref = c
            break
    if col_ref is None:
        cand = [(d[c].astype(str).str.len().mean(), c)
                for c in d.columns if c != col_url]
        cand = [t for t in cand if t[0] >= 4]
        if cand:
            col_ref = max(cand)[1]
    if col_ref is None:
        raise ValueError(
            "No encuentro la columna de referencia.\n"
            "  Columnas del fichero: %s" % list(d.columns))

    # Columna del proveedor. Solo se usa para armar las sentencias SQL, asi
    # que si no aparece no se aborta: simplemente no se generan.
    col_sup = None
    for c in d.columns:
        if c.upper().replace(" ", "_") in NOMBRES_SUPPLIER:
            col_sup = c
            break

    d[col_url] = d[col_url].map(limpiar_url)
    return d, col_ref, col_url, col_sup


# ===================================================================
#  CACHE DE FOTOS
# ===================================================================
def nombre_cache(url):
    """
    El nombre del fichero de cache sale de la URL, no de la referencia, para
    que dos productos con la misma referencia no se pisen la foto.
    """
    return hashlib.sha1(str(url).encode("utf-8")).hexdigest() + ".jpg"


def obtener_imagen(url):
    """
    Devuelve la foto del producto, usando la copia local si ya existe.

    Detalle importante: cuando hay que descargarla, se guarda y se vuelve a
    LEER del disco antes de analizarla, en lugar de analizar la imagen que
    acaba de llegar. Parece un rodeo, pero es lo que hace que el resultado
    sea siempre el mismo.

    Guardar en jpeg con calidad 88 cambia minimamente los pixeles. Si la
    primera pasada analizara la imagen original y las siguientes la copia
    guardada, el color podria salir distinto entre una pasada y otra.
    Medido: de 120 productos, 8 cambiaban de nombre de color y uno pasaba
    de "Cream" a "Black" (era un producto de dos colores). Con dos
    personas procesando el mismo listado, cada una obtendria una cosa.

    Leyendo siempre del disco, el resultado es identico siempre. Y coincide
    con las condiciones en las que se midio la precision del programa, que
    tambien se hizo sobre copias guardadas a calidad 88.
    """
    p = os.path.join(CACHE, nombre_cache(url))
    if os.path.exists(p) and os.path.getsize(p) > 1000:
        try:
            return Image.open(p).convert("RGB")
        except Exception:
            pass   # copia corrupta o a medio escribir: se vuelve a bajar

    img = descargar_imagen(url)
    try:
        img.save(p, quality=88)
        return Image.open(p).convert("RGB")
    except Exception:
        # Si no se puede guardar (carpeta sin permiso, disco lleno), se
        # analiza la imagen descargada y se sigue adelante.
        return img


def procesar_fila(datos):
    pos, url = datos
    if not url.startswith(("http://", "https://")):
        return pos, {"color": "ERROR_URL", "familia": "", "confianza": 0.0,
                     "rgb": None, "revisar": True,
                     "motivo": "la fila no tiene un enlace de foto valido"}
    try:
        return pos, detectar_color(obtener_imagen(url))
    except Exception as e:
        return pos, {"color": "ERROR_DESCARGA", "familia": "", "confianza": 0.0,
                     "rgb": None, "revisar": True,
                     "motivo": "no se pudo obtener la foto: " + str(e)[:80]}


# ===================================================================
#  PROCESO DE UN LISTADO
# ===================================================================
def procesar_listado(entrada, salida=None, hilos=8, tabla=None, sql=False):
    entrada = os.path.abspath(entrada)
    d, col_ref, col_url, col_sup = leer_listado(entrada)
    # Sin --sql, se hace como si no hubiera columna de proveedor: asi se
    # salta todo lo de las sentencias (columna, hojas, avisos y resumen).
    if not sql:
        col_sup = None
    n = len(d)
    if salida is None:
        salida = os.path.join(
            os.path.dirname(entrada),
            os.path.splitext(os.path.basename(entrada))[0] + " CON COLOR.xlsx")
    salida = os.path.abspath(salida)

    print("  Entrada    : %s" % entrada)
    print("  Salida     : %s" % salida)
    print("  Productos  : %d" % n)
    if sql:
        print("  Detectado  : referencia='%s'   fotos='%s'   proveedor='%s'"
              % (col_ref, col_url, col_sup if col_sup else "no encontrada"))
    else:
        print("  Detectado  : referencia='%s'   fotos='%s'" % (col_ref, col_url))
    if sql and col_sup is None:
        print("  Aviso      : sin columna de proveedor no se pueden generar")
        print("               las sentencias SQL. El resto funciona igual.")

    os.makedirs(CACHE, exist_ok=True)

    # Los resultados se acumulan en listas y se pegan como columnas al final.
    # Escribiendolos celda a celda, pandas rechaza el decimal de 'Confianza'
    # cuando la tabla se ha leido como texto.
    color = [""] * n
    familia = [""] * n
    confianza = [None] * n
    revisar = [""] * n
    motivo = [""] * n
    rgb_hex = [""] * n

    # Reanudar: si ya existe la salida, se conserva lo ya resuelto
    hechos = {}
    if os.path.exists(salida):
        try:
            ant = pd.read_excel(salida)
            if col_ref in ant.columns and "Color" in ant.columns:
                for _, f in ant.iterrows():
                    c = str(f.get("Color") or "")
                    if c and c != "nan" and not c.startswith("ERROR"):
                        hechos[str(f[col_ref])] = f
                if hechos:
                    print("  Reanudando : se conservan %d filas ya resueltas"
                          % len(hechos))
        except Exception:
            print("  Aviso      : no pude leer la salida anterior; recalculo todo")

    pendientes = []
    for pos in range(n):
        ref = str(d.at[pos, col_ref])
        if ref in hechos:
            f = hechos[ref]
            color[pos] = str(f.get("Color") or "")
            familia[pos] = str(f.get("Familia") or "")
            try:
                confianza[pos] = float(f.get("Confianza"))
            except (TypeError, ValueError):
                confianza[pos] = None
            revisar[pos] = str(f.get("Revisar") or "")
            v = f.get("Motivo")
            motivo[pos] = "" if v is None or pd.isna(v) else str(v)
            v = f.get("RGB")
            rgb_hex[pos] = "" if v is None or pd.isna(v) else str(v)
        else:
            pendientes.append((pos, str(d.at[pos, col_url])))

    ya = len(os.listdir(CACHE)) if os.path.isdir(CACHE) else 0
    print("  A procesar : %d   (fotos guardadas en local: %d)"
          % (len(pendientes), ya))
    print("  Cache      : %s" % CACHE)
    print()

    t0 = time.time()
    hecho = 0
    if pendientes:
        with ThreadPoolExecutor(max_workers=max(1, hilos)) as ex:
            for pos, r in ex.map(procesar_fila, pendientes):
                color[pos] = r.get("color") or ""
                familia[pos] = r.get("familia") or ""
                confianza[pos] = r.get("confianza")
                revisar[pos] = "SI" if r.get("revisar") else "NO"
                motivo[pos] = r.get("motivo") or ""
                t = r.get("rgb")
                rgb_hex[pos] = ("#%02X%02X%02X" % t) if t else ""
                hecho += 1
                if hecho % 100 == 0 or hecho == len(pendientes):
                    seg = time.time() - t0
                    resto = (seg / hecho) * (len(pendientes) - hecho)
                    print("    %d/%d   %.0f s transcurridos, quedan ~%.0f s"
                          % (hecho, len(pendientes), seg, resto), flush=True)

    if col_url.startswith("Unnamed"):
        d = d.rename(columns={col_url: "URL_Imagen"})

    d["Color"] = color
    d["Familia"] = familia
    d["Confianza"] = confianza
    d["Revisar"] = revisar
    d["Motivo"] = motivo
    d["RGB"] = rgb_hex

    # ---- Sentencias UPDATE, una por producto -------------------------
    # Se separan en dos listas: las de color fiable y las de color dudoso,
    # para no ejecutar a ciegas un color que puede estar mal. Las filas sin
    # color (foto caida) no generan sentencia: no habria nada que escribir.
    sql_col = [""] * n
    sql_listo, sql_dudoso = [], []
    if col_sup is not None:
        for pos in range(n):
            c = str(color[pos] or "")
            if not c or c.startswith("ERROR"):
                continue
            s = sentencia_update(c, d.at[pos, col_sup], d.at[pos, col_ref],
                                 tabla)
            sql_col[pos] = s
            (sql_dudoso if revisar[pos] == "SI" else sql_listo).append(s)
        d["SQL"] = sql_col

    try:
        with pd.ExcelWriter(salida, engine="openpyxl") as libro:
            d.to_excel(libro, sheet_name="Productos", index=False)
            if col_sup is not None:
                pd.DataFrame({"-- sentencias listas para ejecutar": sql_listo}) \
                    .to_excel(libro, sheet_name="SQL listo", index=False)
                pd.DataFrame({"-- revisar el color antes de ejecutar": sql_dudoso}) \
                    .to_excel(libro, sheet_name="SQL a revisar", index=False)
                # Ensanchar la columna del SQL para que se lea sin pelearse
                for hoja in ("SQL listo", "SQL a revisar"):
                    libro.sheets[hoja].column_dimensions["A"].width = 140
    except PermissionError:
        raise ValueError(
            "No puedo escribir el resultado porque el fichero esta abierto:\n"
            "  %s\n  Cierralo en Excel y vuelve a lanzarlo." % salida)

    n_rev = int((d["Revisar"] == "SI").sum())
    n_err = int(d["Color"].astype(str).str.startswith("ERROR").sum())
    print()
    print("  LISTO en %.0f s" % (time.time() - t0))
    print("  Con color asignado    : %d de %d" % (n - n_err, n))
    if n_err:
        print("  Sin foto / con error  : %d" % n_err)
    print("  Marcados 'Revisar=SI' : %d  (%.1f%%)"
          % (n_rev, 100.0 * n_rev / n if n else 0))
    print()
    print("  Colores mas frecuentes:")
    for nombre, veces in d["Color"].value_counts().head(8).items():
        print("     %-14s %d" % (nombre, veces))
    if col_sup is not None:
        print()
        print("  Sentencias SQL generadas:")
        print("     hoja 'SQL listo'     : %d  (color fiable)" % len(sql_listo))
        print("     hoja 'SQL a revisar' : %d  (mirar la foto antes)"
              % len(sql_dudoso))
    print()
    print("  >>> FICHERO GENERADO: %s" % salida)
    return salida, n, n_rev, n_err


# ===================================================================
#  ARRANQUE
# ===================================================================
def elegir_ficheros():
    """Abre un selector de ficheros. Se usa cuando no se indica ninguno."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return []
    raiz = tk.Tk()
    raiz.withdraw()
    inicio = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(inicio):
        inicio = os.path.expanduser("~")
    rutas = filedialog.askopenfilenames(
        title="Elige el listado de productos (se pueden elegir varios)",
        initialdir=inicio,
        filetypes=[("Listados", "*.csv *.xlsx *.xls *.xlsm"),
                   ("CSV", "*.csv"), ("Excel", "*.xlsx *.xls *.xlsm"),
                   ("Todos", "*.*")])
    raiz.destroy()
    return list(rutas)


def main():
    ap = argparse.ArgumentParser(
        description="Detecta el color de cada producto analizando su foto. "
                    "Todo en local: sin API ni claves.")
    ap.add_argument("entradas", nargs="*",
                    help="uno o varios .csv/.xlsx. Si no pones ninguno, se "
                         "abre un selector de ficheros.")
    ap.add_argument("-s", "--salida", default=None,
                    help="fichero .xlsx de salida (solo si procesas UN listado)")
    ap.add_argument("-w", "--hilos", type=int, default=8,
                    help="descargas en paralelo (por defecto 8)")
    ap.add_argument("--sql", action="store_true",
                    help="genera tambien las sentencias SQL UPDATE (hace falta "
                         "una columna de proveedor en el listado)")
    ap.add_argument("-t", "--tabla", default=TABLA_SQL,
                    help="tabla para las sentencias UPDATE, solo con --sql "
                         "(por defecto %s)" % TABLA_SQL)
    args = ap.parse_args()

    entradas = [e for e in args.entradas if e.strip()]
    if not entradas:
        print("No has indicado ningun fichero: abro el selector...")
        entradas = elegir_ficheros()
    if not entradas:
        print()
        print("No has elegido ningun fichero. No hay nada que hacer.")
        print()
        print("Formas de usarlo:")
        print("  - Arrastra el listado encima de 'ETIQUETAR COLORES.bat'")
        print("  - O haz doble clic en 'ETIQUETAR COLORES.bat' y eligelo")
        return 0

    if args.salida and len(entradas) > 1:
        print("Aviso: --salida solo vale con un listado. Se ignora.")
        args.salida = None

    fallos = 0
    for i, ruta in enumerate(entradas, 1):
        print()
        print("=" * 68)
        print("LISTADO %d de %d: %s" % (i, len(entradas), os.path.basename(ruta)))
        print("=" * 68)
        try:
            procesar_listado(ruta, args.salida if len(entradas) == 1 else None,
                             args.hilos, args.tabla, args.sql)
        except ValueError as e:
            fallos += 1
            print()
            print("  NO SE HA PODIDO PROCESAR")
            print("  %s" % str(e).replace("\n", "\n  "))
        except Exception as e:
            fallos += 1
            print()
            print("  ERROR INESPERADO: %s: %s" % (type(e).__name__, e))
            print("  Detalle tecnico:")
            print("  " + traceback.format_exc().replace("\n", "\n  "))

    print()
    print("=" * 68)
    if fallos:
        print("Terminado con %d listado(s) que no se han podido procesar." % fallos)
    else:
        print("Terminado. Abre el Excel y filtra la columna 'Revisar' por 'SI'")
        print("para ver solo las filas que conviene mirar a ojo.")
    print("=" * 68)
    return 1 if fallos else 0


if __name__ == "__main__":
    sys.exit(main())
