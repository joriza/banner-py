"""Generador determinista del cartel de letras huecas para recortar.

Complemento de cartelería del taller de trípticos: produce un DOCX
compatible con Word 2010, hoja A4 apaisada, con 2 letras por hoja
(una al lado de la otra). Cada letra sale "hueca" (efecto de fuente
Contorno de Word, <w:outline/>) para que los alumnos la coloreen a
su gusto y luego la recorten siguiendo la guía punteada.

El texto se selecciona manualmente en la línea de comandos:

    python generar-banner-letras.py [TEXTO]

Sin argumento genera "CENS460". Los espacios separan palabras y no
consumen hoja: solo se imprimen las letras.

Única dependencia extra: Pillow (PIL), para medir los glifos y
calcular el tamaño de letra más grande que entra en cada media hoja.
Se mide con Arial Bold (métrica proxy) y un factor de seguridad, y
el DOCX nombra "Arial Black"; si la máquina destino no la tiene,
Word sustituye la fuente y la letra sigue entrando.

Exit code: 0 si todo salió bien; 2 argumento inválido; 3 datos,
entorno o medición inválidos; 4 determinismo roto; 5 XML mal
formado; 6 error de escritura.

Reglas de estilo de este archivo (linter, igual que el generador de
trípticos):
    - los imports van dentro de las funciones, nunca al nivel del módulo;
    - nada de statements compuestos en una línea ni de ``;``;
    - try/except alrededor de toda I/O de archivos y del parseo XML;
    - nada de ``assert``: chequeos explícitos con mensaje y exit code.
"""

FECHA_FIJA_ZIP = (1980, 1, 1, 0, 0, 0)

TEXTO_POR_DEFECTO = "CENS460"

# ---------------------------------------------------------------- geometría de página
# A4 apaisada en twips (1 cm = 566.929 twips).
PG_W = 16838
PG_H = 11906
MARGEN_SUP = 284
MARGEN_INF = 567
MARGEN_IZQ = 284
MARGEN_DER = 284
PIE_DIST = 227

ANCHO_UTIL = PG_W - MARGEN_IZQ - MARGEN_DER
ALTO_UTIL = PG_H - MARGEN_SUP - MARGEN_INF
ANCHO_COLUMNA = ANCHO_UTIL // 2
ALTO_FILA = ALTO_UTIL - 35

# Márgenes de aire alrededor de la tinta de cada letra, en cm,
# para que la tijera entre cómoda.
PAD_LATERAL_CM = 0.35
PAD_SUPERIOR_CM = 0.30

# Factores de seguridad de la medición con la fuente proxy:
# Arial Black es ~10% más ancha que Arial Bold y la negrita simulada
# agrega un poco más; la altura de mayúsculas es casi idéntica.
FACTOR_ANCHO = 1.14
FACTOR_ALTO = 1.03

PT_POR_CM = 28.3465
TWIPS_POR_CM = 566.929

FUENTES_PROXY = (
    "C:/Windows/Fonts/arialbd.ttf",
    "C:/Windows/Fonts/arial.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
    "/Library/Fonts/Arial Bold.ttf",
)

FUENTE_DOCX = "Arial Black"
COLOR_GUIA = "BFBFBF"
COLOR_PIE = "808080"


class ErrorGeneracion(Exception):
    """Falla con mensaje para el usuario y exit code propio."""

    def __init__(self, mensaje, codigo):
        super().__init__(mensaje)
        self.codigo = codigo


# ---------------------------------------------------------------- paquete OOXML

CABECERA_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

CONTENT_TYPES_XML = (
    CABECERA_XML
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    + "</Types>"
)

RELS_RAIZ_XML = (
    CABECERA_XML
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + "</Relationships>"
)

RELS_DOCUMENTO_XML = (
    CABECERA_XML
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    + "</Relationships>"
)

STYLES_XML = (
    CABECERA_XML
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + "<w:docDefaults><w:rPrDefault><w:rPr>"
    + '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    + '<w:sz w:val="22"/><w:szCs w:val="22"/>'
    + '<w:lang w:val="es-AR" w:eastAsia="es-AR" w:bidi="ar-SA"/>'
    + "</w:rPr></w:rPrDefault>"
    + '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    + "</w:docDefaults>"
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    + "</w:styles>"
)


def _campo_pie(instruccion, texto_previo):
    """Campo simple de Word (PAGE o NUMPAGES) con el estilo del pie."""
    from xml.sax.saxutils import escape

    rpr = (
        '<w:rPr><w:color w:val="' + COLOR_PIE + '"/>'
        + '<w:sz w:val="14"/><w:szCs w:val="14"/></w:rPr>'
    )
    return (
        '<w:fldSimple w:instr=" ' + instruccion + ' ">'
        + "<w:r>" + rpr + "<w:t>" + escape(texto_previo) + "</w:t></w:r>"
        + "</w:fldSimple>"
    )


def _pie_xml(texto):
    """Pie de página: identifica el cartel y numera las hojas."""
    from xml.sax.saxutils import escape

    rpr = (
        '<w:rPr><w:color w:val="' + COLOR_PIE + '"/>'
        + '<w:sz w:val="14"/><w:szCs w:val="14"/></w:rPr>'
    )
    return (
        CABECERA_XML
        + '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/><w:jc w:val=\"center\"/></w:pPr>"
        + "<w:r>" + rpr + "<w:t xml:space=\"preserve\">"
        + escape(texto + " · Página ")
        + "</w:t></w:r>"
        + _campo_pie("PAGE", "1")
        + "<w:r>" + rpr + "<w:t xml:space=\"preserve\"> de </w:t></w:r>"
        + _campo_pie("NUMPAGES", "1")
        + "</w:p></w:ftr>"
    )


def _letra_run_xml(caracter, sz_half):
    """Run de una letra hueca: contorno negro sin relleno (efecto Contorno)."""
    from xml.sax.saxutils import escape

    rpr = (
        "<w:rPr>"
        + '<w:rFonts w:ascii="' + FUENTE_DOCX + '" w:hAnsi="' + FUENTE_DOCX + '" w:cs="' + FUENTE_DOCX + '"/>'
        + "<w:b/><w:bCs/>"
        + "<w:outline/>"
        + '<w:color w:val="000000"/>'
        + '<w:sz w:val="' + str(sz_half) + '"/><w:szCs w:val="' + str(sz_half) + '"/>'
        + "</w:rPr>"
    )
    return "<w:r>" + rpr + '<w:t xml:space="preserve">' + escape(caracter) + "</w:t></w:r>"


def _celda_xml(caracter, sz_half):
    """Media hoja: una letra centrada, o vacía si sobra la mitad."""
    if caracter is None:
        parrafo = (
            "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/>"
            + '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
        )
    else:
        parrafo = (
            "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" w:line=\""
            + str(ALTO_FILA)
            + "\" w:lineRule=\"exact\"/><w:jc w:val=\"center\"/>"
            + '<w:rPr><w:sz w:val="' + str(sz_half) + '"/></w:rPr></w:pPr>'
            + _letra_run_xml(caracter, sz_half)
            + "</w:p>"
        )
    return (
        "<w:tc><w:tcPr>"
        + '<w:tcW w:w="' + str(ANCHO_COLUMNA) + '" w:type="dxa"/>'
        + '<w:vAlign w:val="center"/>'
        + "</w:tcPr>"
        + parrafo
        + "</w:tc>"
    )


def _documento_xml(texto, sz_half):
    """Cuerpo: una tabla de 2 columnas; cada fila exacta de una página."""
    caracteres = [c for c in texto if not c.isspace()]
    filas = []
    for indice in range(0, len(caracteres), 2):
        izquierda = caracteres[indice]
        derecha = caracteres[indice + 1] if indice + 1 < len(caracteres) else None
        filas.append(
            "<w:tr><w:trPr>"
            + '<w:trHeight w:val="' + str(ALTO_FILA) + '" w:hRule="exact"/>'
            + "<w:cantSplit/>"
            + "</w:trPr>"
            + _celda_xml(izquierda, sz_half)
            + _celda_xml(derecha, sz_half)
            + "</w:tr>"
        )

    tabla = (
        "<w:tbl><w:tblPr>"
        + '<w:tblW w:w="' + str(ANCHO_UTIL) + '" w:type="dxa"/>'
        + '<w:tblLayout w:type="fixed"/>'
        + "<w:tblBorders>"
        + '<w:top w:val="dashed" w:sz="4" w:space="0" w:color="' + COLOR_GUIA + '"/>'
        + '<w:left w:val="dashed" w:sz="4" w:space="0" w:color="' + COLOR_GUIA + '"/>'
        + '<w:bottom w:val="dashed" w:sz="4" w:space="0" w:color="' + COLOR_GUIA + '"/>'
        + '<w:right w:val="dashed" w:sz="4" w:space="0" w:color="' + COLOR_GUIA + '"/>'
        + '<w:insideH w:val="dashed" w:sz="4" w:space="0" w:color="' + COLOR_GUIA + '"/>'
        + '<w:insideV w:val="dashed" w:sz="4" w:space="0" w:color="' + COLOR_GUIA + '"/>'
        + "</w:tblBorders>"
        + '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        + '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        + "</w:tblPr>"
        + '<w:tblGrid><w:gridCol w:w="' + str(ANCHO_COLUMNA) + '"/>'
        + '<w:gridCol w:w="' + str(ANCHO_UTIL - ANCHO_COLUMNA) + '"/></w:tblGrid>'
        + "".join(filas)
        + "</w:tbl>"
        + '<w:p><w:pPr><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
    )

    sectpr = (
        "<w:sectPr>"
        + '<w:footerReference w:type="default" r:id="rId2"/>'
        + '<w:pgSz w:w="' + str(PG_W) + '" w:h="' + str(PG_H) + '" w:orient="landscape"/>'
        + '<w:pgMar w:top="' + str(MARGEN_SUP) + '" w:right="' + str(MARGEN_DER)
        + '" w:bottom="' + str(MARGEN_INF) + '" w:left="' + str(MARGEN_IZQ)
        + '" w:header="284" w:footer="' + str(PIE_DIST) + '" w:gutter="0"/>'
        + '<w:cols w:space="708"/>'
        + '<w:docGrid w:linePitch="360"/>'
        + "</w:sectPr>"
    )

    return (
        CABECERA_XML
        + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + "<w:body>"
        + tabla
        + sectpr
        + "</w:body></w:document>"
    )


def _partes_docx(texto, sz_half, pie):
    """Las 6 partes del paquete, en orden fijo."""
    return (
        ("[Content_Types].xml", CONTENT_TYPES_XML),
        ("_rels/.rels", RELS_RAIZ_XML),
        ("word/document.xml", _documento_xml(texto, sz_half)),
        ("word/styles.xml", STYLES_XML),
        ("word/footer1.xml", _pie_xml(pie)),
        ("word/_rels/document.xml.rels", RELS_DOCUMENTO_XML),
    )


def _docx_bytes(texto, sz_half, pie):
    import io
    import zipfile

    buffer = io.BytesIO()
    try:
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as paquete:
            for nombre, contenido in _partes_docx(texto, sz_half, pie):
                entrada = zipfile.ZipInfo(nombre, date_time=FECHA_FIJA_ZIP)
                entrada.compress_type = zipfile.ZIP_DEFLATED
                entrada.create_system = 0
                entrada.external_attr = 0o600 << 16
                paquete.writestr(entrada, contenido.encode("utf-8"), compresslevel=9)
    except ErrorGeneracion:
        raise
    except Exception as exc:
        raise ErrorGeneracion("No pude empaquetar el DOCX: " + repr(exc), 3) from exc
    return buffer.getvalue()


# ---------------------------------------------------------------- medición del tamaño

def _ruta_fuente_proxy():
    import os

    for ruta in FUENTES_PROXY:
        if os.path.exists(ruta):
            return ruta
    raise ErrorGeneracion(
        "No encontré una fuente Arial para medir. Probé: " + ", ".join(FUENTES_PROXY),
        3,
    )


def _medir_tamano_pt(texto):
    """Mayor tamaño de fuente (pt) en que cada letra entra en su media hoja.

    Devuelve (tamano_entero, descripcion_del_limite). Mide la tinta de
    cada carácter con la fuente proxy y aplica los factores de seguridad
    sobre el ancho y el alto disponibles de la celda.
    """
    from PIL import ImageFont

    ruta = _ruta_fuente_proxy()
    try:
        fuente = ImageFont.truetype(ruta, 100)
    except Exception as exc:
        raise ErrorGeneracion(
            "No pude cargar la fuente proxy " + ruta + ": " + repr(exc), 3
        ) from exc

    ancho_max_pt = (ANCHO_COLUMNA / TWIPS_POR_CM - PAD_LATERAL_CM) * PT_POR_CM
    alto_max_pt = (ALTO_FILA / TWIPS_POR_CM - PAD_SUPERIOR_CM) * PT_POR_CM

    caracteres = []
    for caracter in texto:
        if not caracter.isspace() and caracter not in caracteres:
            caracteres.append(caracter)
    if len(caracteres) == 0:
        raise ErrorGeneracion("El texto no tiene ninguna letra para imprimir.", 3)

    tamano = None
    limite = ""
    for caracter in caracteres:
        try:
            x0, y0, x1, y1 = fuente.getbbox(caracter)
        except Exception as exc:
            raise ErrorGeneracion(
                "No pude medir el carácter " + repr(caracter) + ": " + repr(exc), 3
            ) from exc
        ancho_tinta = (x1 - x0) * FACTOR_ANCHO
        alto_tinta = (y1 - y0) * FACTOR_ALTO
        if ancho_tinta <= 0 or alto_tinta <= 0:
            raise ErrorGeneracion(
                "El carácter " + repr(caracter) + " no tiene tinta medible.", 3
            )
        por_ancho = ancho_max_pt / ancho_tinta * 100.0
        por_alto = alto_max_pt / alto_tinta * 100.0
        candidato = min(por_ancho, por_alto)
        if tamano is None or candidato < tamano:
            tamano = candidato
            if por_ancho <= por_alto:
                limite = "ancho de la media hoja"
            else:
                limite = "alto de la media hoja"

    if tamano is None:
        raise ErrorGeneracion(
            "No pude medir ningún carácter del texto: " + repr(texto), 3
        )
    tamano_entero = int(tamano)
    if tamano_entero < 40:
        raise ErrorGeneracion(
            "El texto es demasiado largo: la letra quedaría de menos de 40 pt.", 3
        )
    return tamano_entero, limite, caracteres


# ---------------------------------------------------------------- validación y escritura

def _normalizar_texto(bruto):
    """Mayúsculas, sin espacios en los extremos; valida el contenido."""
    texto = bruto.strip().upper()
    if len(texto) == 0:
        raise ErrorGeneracion("El texto del cartel está vacío.", 3)
    if len(texto) > 16:
        raise ErrorGeneracion(
            "Texto demasiado largo (" + str(len(texto)) + " caracteres, máximo 16).", 3
        )
    for caracter in texto:
        if caracter.isspace():
            continue
        if ord(caracter) < 33 or ord(caracter) > 0x17F:
            raise ErrorGeneracion(
                "Carácter no imprimible en el cartel: " + repr(caracter), 3
            )
    return texto


def _nombre_archivo(texto):
    import re
    import unicodedata

    descompuesto = unicodedata.normalize("NFKD", texto)
    plano = "".join(c for c in descompuesto if not unicodedata.combining(c))
    slug = re.sub(r"[^A-Z0-9]+", "-", plano).strip("-")
    if len(slug) == 0:
        slug = "cartel"
    return "banner-" + slug.lower() + "-letras-huecas.docx"


def _verificar_xml(texto, sz_half, pie):
    from defusedxml import ElementTree

    for nombre_parte, contenido in _partes_docx(texto, sz_half, pie):
        try:
            ElementTree.fromstring(contenido.encode("utf-8"))
        except Exception as exc:
            raise ErrorGeneracion(
                "XML mal formado :: " + nombre_parte + ": " + repr(exc), 5
            ) from exc


def _escribir_docx(ruta_destino, datos):
    try:
        with open(ruta_destino, "wb") as salida:
            salida.write(datos)
    except OSError as exc:
        raise ErrorGeneracion(
            "No pude escribir " + ruta_destino + ": " + repr(exc), 6
        ) from exc


def _generar(texto, carpeta_destino):
    import math
    import os

    tamano_pt, limite, caracteres = _medir_tamano_pt(texto)
    sz_half = tamano_pt * 2
    pie = texto + " · letras huecas para recortar y colorear"

    corrida_a = _docx_bytes(texto, sz_half, pie)
    corrida_b = _docx_bytes(texto, sz_half, pie)
    if corrida_a != corrida_b:
        raise ErrorGeneracion(
            "Determinismo roto: dos corridas del mismo cartel difieren ("
            + str(len(corrida_a))
            + " vs "
            + str(len(corrida_b))
            + " bytes).",
            4,
        )
    _verificar_xml(texto, sz_half, pie)

    ruta_destino = os.path.join(carpeta_destino, _nombre_archivo(texto))
    _escribir_docx(ruta_destino, corrida_a)

    hojas = int(math.ceil(len(caracteres) / 2.0))
    return {
        "ruta": ruta_destino,
        "bytes": len(corrida_a),
        "sz_pt": tamano_pt,
        "limite": limite,
        "hojas": hojas,
        "letras": len(caracteres),
    }


# ---------------------------------------------------------------- CLI


def main(argumentos):
    import os
    import sys

    if len(argumentos) > 1:
        print(
            "Uso: python generar-banner-letras.py [TEXTO]",
            file=sys.stderr,
        )
        return 2
    if len(argumentos) == 1:
        bruto = argumentos[0]
    else:
        bruto = TEXTO_POR_DEFECTO
    carpeta_destino = os.path.dirname(os.path.abspath(__file__))
    try:
        texto = _normalizar_texto(bruto)
        resultado = _generar(texto, carpeta_destino)
    except ErrorGeneracion as exc:
        print("ERROR: " + str(exc), file=sys.stderr)
        return exc.codigo
    print(
        "OK "
        + os.path.basename(resultado["ruta"])
        + " ("
        + str(resultado["bytes"])
        + " bytes, determinista, XML bien formado)"
    )
    print(
        "Texto: "
        + texto
        + " · "
        + str(resultado["letras"])
        + " letras en "
        + str(resultado["hojas"])
        + " hoja(s) A4 apaisada(s), 2 letras por hoja"
    )
    print(
        "Tamaño de letra: "
        + str(int(resultado["sz_pt"]))
        + " pt (limitado por el "
        + resultado["limite"]
        + "), fuente "
        + FUENTE_DOCX
        + " hueca"
    )
    return 0


if __name__ == "__main__":
    import sys

    codigo_salida = main(sys.argv[1:])
    sys.exit(codigo_salida)
