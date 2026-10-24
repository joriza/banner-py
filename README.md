# banner-py

Generador de carteles con letras grandes huecas para recortar y colorear, pensado para el aula.
Genera documentos **Word 2010 (.docx)** listos para imprimir: cada hoja trae letras gigantes con
guías punteadas de recorte, que los alumnos colorean y luego arman como cartel.

## Descargar el programa portable

El ejecutable para Windows está en [Releases](https://github.com/joriza/banner-py/releases).
Es un único archivo: se copia al pendrive y se ejecuta con doble clic. No instala nada y no
necesita Python ni Word para generar.

> Requiere Windows 8.1 o superior. Para abrir e imprimir el resultado: Word 2010 o compatible.

## Uso

Doble clic y responder (Enter toma el valor entre corchetes):

```
Texto del cartel [CENS460]:
Orientacion: [V]ertical o [A]paisada [A]:
Hoja: [1] A4  [2] A5  [3] Letter  [4] Medida propia [1]:
Letras por hoja (1 a 12) [2]:
Fuente: [1] Arial Black [2] Impact [3] Arial [4] Rockwell [5] Bookman [1]:
Estilo: [H]ueca para colorear o [R]ellena negra [H]:
Espacios: [Q]uitar (letras sueltas) o [M]antener (cartel continuo) [Q]:
```

O por argumentos:

```
banner-portable.exe CENS460
banner-portable.exe HOLA --vertical --letras 3 --fuente impact --hoja a5
banner-portable.exe CARTEL --ancho 30 --alto 21 --letras 4
banner-portable.exe CARTA --rellena --espacios mantener --fuente rockwell
```

## Qué hace

- **Letra hueca o rellena**: contorno negro para colorear, o negro macizo.
- **Grilla automática óptima**: con N letras por hoja prueba todas las combinaciones
  columna × fila y elige la de letra más grande.
- **Orden de lectura**: izquierda a derecha, arriba a abajo. El resumen de consola
  indica qué letra va en cada hoja.
- **Tamaño máximo medido**: se mide la tinta real de cada glifo con el motor de texto
  de Windows (GDI). Si la fuente no está instalada, mide la sustituta con margen de
  seguridad y, como último recurso, usa una tabla de métricas embebida: nunca se
  queda sin poder medir.
- **Hoja a medida**: A4, A5, Letter o ancho × alto libre en cm.
- **No pisa archivos**: si la salida ya existe, agrega `_V1`, `_V2`, … con el primer
  número libre (también con `-o` explícita).
- **Determinista**: mismos datos y misma máquina producen exactamente el mismo archivo.
- **Fuente**: Arial Black, Impact y Arial vienen con Windows; Rockwell y Bookman Old
  Style vienen con Office. Si una falta, se usa la sustituta y la letra sale apenas
  más chica, sin romper nada.

## Archivos del repo

| Archivo | Qué es |
| --- | --- |
| `banner-portable.py` | Fuente del programa portable |
| `compilar-banner-exe.bat` | Recompila el exe con PyInstaller |

El ejecutable **no** se versiona en el repo: se publica en
[Releases](https://github.com/joriza/banner-py/releases).
Los archivos generados (`.pdf`, `.doc`, `.docx`) también están excluidos, salvo los
que se llamen `example`.

## Compilar el exe

```
pip install pyinstaller
compilar-banner-exe.bat
```

## Detalles técnicos

- El DOCX se genera como OOXML crudo, sin dependencias: tabla con filas de altura
  exacta y letra con el efecto Contorno de Word (`<w:outline/>`).
- Compatible con Word 2010 (modo de compatibilidad nativo, sin repair prompts).
- Pie de página con numeración `Página X de Y` para ordenar las hojas al armar el cartel.
