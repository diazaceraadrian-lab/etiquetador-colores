# Etiquetador de colores

Herramienta para catálogos de moda: le das un listado de productos con la URL
de la foto de cada uno y te devuelve el mismo listado en Excel con el **color
de cada producto**, listo para filtrar o para cargarlo en la base de datos.

Todo se calcula en tu ordenador mirando los píxeles de cada foto. **No usa
servicios de IA de pago ni claves de API**: coste cero y nada que se pueda
filtrar.

La hice para un caso real: poner a mano el color de cada producto de los
listados de los proveedores, uno a uno.

## Resultados medidos

Validado contra **237 fotos de dos proveedores**:

| | |
|---|---|
| Acierta la familia de color exacta | **81 %** |
| Acierta la exacta o una muy parecida | **97 %** |
| Fallos claros | 3 % (7 de 237) |
| Velocidad | ~1.000 productos en 3,5 minutos (40 s si las fotos ya están en caché) |

"Muy parecida" significa confundir tonos que a la vista se parecen de verdad:
beige con marrón, gris con negro, rosa con morado.

## Qué añade al listado

| Columna | Qué es |
|---|---|
| `Color` | Nombre comercial en inglés: *Black*, *Navy blue*, *Camel*… |
| `Familia` | Grupo de color (*black*, *blue*, *brown*…), para filtrar |
| `Confianza` | De 0 a 1: qué parte del producto es de ese color |
| `Revisar` | `SI` / `NO`: si conviene que lo mire una persona |
| `Motivo` | Por qué se marca para revisar |
| `RGB` | El color medido, en formato `#RRGGBB` |

Solo hace falta revisar a mano las filas con `Revisar = SI` (normalmente entre
el 8 % y el 17 %). Opcionalmente, también genera las sentencias SQL `UPDATE`
listas para cargar los colores en la base de datos.

## Cómo funciona por dentro

1. **Descarga la foto** y la reduce a 420 px (guarda una copia en caché para
   no volver a descargarla).
2. **Separa el producto del fondo** blanco de estudio: el fondo es la zona
   clara y sin color conectada con el borde de la foto. Si sale un modelo,
   descarta también los píxeles de piel.
3. **Agrupa los colores del producto** con *k-means* (un algoritmo que reparte
   los píxeles en unos pocos grupos de colores parecidos) y se queda con el
   dominante.
4. **Lo traduce al nombre comercial más cercano** midiendo la distancia en el
   espacio de color CIE Lab, donde la distancia entre dos colores se parece a
   la diferencia que percibe el ojo humano.

## Instalación y uso

Necesitas Python 3.

```bash
pip install -r requirements.txt
python etiquetar_colores_local.py mis_productos.csv
```

En Windows también puedes **arrastrar el listado (.csv o .xlsx) encima de
`ETIQUETAR COLORES.bat`**. El resultado aparece al lado, con el nombre
`mis_productos CON COLOR.xlsx`.

Opciones:

| Opción | Para qué |
|---|---|
| `-s`, `--salida` | Nombre del Excel de salida (si procesas un solo listado) |
| `-w`, `--hilos` | Descargas en paralelo (por defecto 8) |
| `--sql` | Genera también las sentencias SQL (ver abajo) |
| `-t`, `--tabla` | Tabla para las sentencias SQL (por defecto, la variable de entorno `ETIQUETADOR_TABLA` o `productos`) |

### Sentencias SQL (opcional)

Están desactivadas por defecto. Si lanzas el programa con `--sql` y el listado
trae una columna de proveedor, el Excel añade una columna `SQL` y dos hojas con
sentencias `UPDATE` (las de color fiable y las que conviene revisar antes):

```bash
python etiquetar_colores_local.py mis_productos.csv --sql
```

```sql
update productos set color = 'Navy blue' where proveedor_id = 7 and referencia = 'AB123';
```

Los nombres de la tabla y de las columnas se adaptan a tu base de datos con
variables de entorno:

| Variable | Por defecto |
|---|---|
| `ETIQUETADOR_TABLA` | `productos` |
| `ETIQUETADOR_COL_COLOR` | `color` |
| `ETIQUETADOR_COL_PROVEEDOR` | `proveedor_id` |
| `ETIQUETADOR_COL_REFERENCIA` | `referencia` |

Encuentra las columnas por su **contenido**, no por su nombre, así que da
igual cómo vengan las cabeceras. Si lo paras a medias, al relanzarlo continúa
donde lo dejó.

## Límites conocidos

- **Dorados y plateados** salen como beige, crema o gris: en una foto, el oro
  y el crema son el mismo color. Se probó a distinguirlos por el brillo y no
  funciona.
- **Vaqueros muy desteñidos** pueden salir grises en vez de azules.
- **Fotos con dos prendas** (un abrigo sobre unos vaqueros): mide el color que
  más ocupa, sin saber cuál es el producto.
- **Productos que no son ropa** (bolígrafos, perfumes) van peor.

La guía de uso pensada para quien no programa está en [`LEEME.txt`](LEEME.txt).
