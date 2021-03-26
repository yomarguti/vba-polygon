# vba-polygon
VBA Macro para calcular la cantidad de puntos que caben sobre una línea circunscrita por un polígono determinado

### cPoint
- x as integer
- y as integer

### Funciones
- **inside_polygon(punto, poligono): boolean**
Calcula si un punto se encuentra dentro del poligono dado.

- **get_points(filename): cPoint()**:
Lee un archivo txt y regresa una lista de puntos

- **get_line_intersection(p0, p1, p2, p3): cPoint**:
Retorna el punto de interseccion de dos segmentos de recta, sino intersecan el punto regresa vacio.

- **apply_regex(patron, string):matches**
Recibe el patron y el string donde aplicar la expresion regular. Devuelve un objeto con las coincidencias despues de ejecuta la expresion regular.

- ** loadFileStr(archivo): content:**
Devuelde el contenido de un archivo de texto

- **arrayLen(cPoint Array): integer**:
Devuelve la longitud de un array de cPoints


