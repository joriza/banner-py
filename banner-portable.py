"""Cartel portable de letras huecas para recortar y colorear.

Programa portable para Windows (un solo ejecutable, sin archivos
externos, pensado para correr desde un pendrive). Genera un DOCX
compatible con Word 2010: hoja elegida (A4 por defecto, o A5,
Letter o medida libre en cm), vertical o apaisada, con N letras
por hoja repartidas de izquierda a derecha y de arriba a abajo.
Cada letra sale "hueca" (efecto Contorno de Word, <w:outline/>),
del mayor tamaño que entra en su celda, para colorear y recortar.

Uso interactivo (doble clic): responde las preguntas y acepta con
Enter para tomar el valor por defecto.

    Texto del cartel [CENS460]:
    Orientacion: [V]ertical o [A]paisada [A]:
    Hoja: [1] A4  [2] A5  [3] Letter  [4] Medida propia [1]:
    Letras por hoja (1 a 12) [2]:
    Fuente: [1] Arial Black [2] Impact [3] Arial [4] Rockwell [5] Bookman [1]:
    Estilo: [H]ueca para colorear o [R]ellena negra [H]:
    Espacios entre palabras: [Q]uitar o [M]antener [Q]:

Uso por argumentos (avanzado):

    banner-portable.exe CENS460
    banner-portable.exe HOLA --vertical --letras 3 --fuente impact --hoja a5
    banner-portable.exe CARTEL --ancho 30 --alto 21 --letras 4
    banner-portable.exe CARTA --rellena --espacios mantener --fuente rockwell

El DOCX se guarda junto al ejecutable (o junto a este archivo .py),
con nombre banner-<texto>-letras-huecas.docx. Nunca pisa un archivo
existente: si el nombre ya está tomado, agrega un sufijo creciente
_V1, _V2, ... con el primer número libre (también con -o explícita).

Determinismo: con los mismos datos y la misma máquina, dos corridas
producen exactamente el mismo archivo (timestamps fijos en el ZIP y
una verificación de doble corrida). Las métricas de letra se miden
con GDI de Windows sobre la fuente real del equipo; en una máquina
sin la fuente se mide la sustituta, y si nada está disponible se usa
una tabla de métricas embebida: el programa nunca se queda sin poder
medir.

Exit code: 0 si todo salió bien; 2 argumento inválido; 3 datos,
entorno o medición inválidos; 4 determinismo roto; 5 XML mal
formado; 6 error de escritura.

Reglas de estilo de este archivo (linter, igual que el resto del
toolchain docente):
    - los imports van dentro de las funciones, nunca al nivel del módulo;
    - nada de statements compuestos en una línea ni de ``;``;
    - try/except alrededor de toda I/O de archivos y del parseo XML;
    - nada de ``assert``: chequeos explícitos con mensaje y exit code.
"""

FECHA_FIJA_ZIP = (1980, 1, 1, 0, 0, 0)

TEXTO_POR_DEFECTO = "CENS460"
TEXTO_MAXIMO = 64
LETRAS_POR_HOJA_MAX = 12
TAMANO_MINIMO_PT = 40

# Presets de hoja: (lado corto cm, lado largo cm).
PRESETS_HOJA = {
    "a4": (21.0, 29.7),
    "a5": (14.8, 21.0),
    "letter": (21.59, 27.94),
}

# Presets de fuente. "archivos": nombres posibles del TTF real en la
# máquina; "embebida": fracciones del cuerpo (em) de último recurso.
FUENTES_PRESET = {
    "arial-black": {
        "nombre": "Arial Black",
        "archivos": ("ariblk.ttf",),
        "embebida": {"base": (0.78, 0.72), "ancha": (0.98, 0.72), "angosta": (0.50, 0.72)},
    },
    "impact": {
        "nombre": "Impact",
        "archivos": ("impact.ttf",),
        "embebida": {"base": (0.64, 0.72), "ancha": (0.88, 0.72), "angosta": (0.42, 0.72)},
    },
    "arial": {
        "nombre": "Arial",
        "archivos": ("arial.ttf", "arialbd.ttf"),
        "embebida": {"base": (0.73, 0.716), "ancha": (0.95, 0.716), "angosta": (0.45, 0.716)},
    },
    "rockwell": {
        "nombre": "Rockwell",
        "archivos": ("rock.ttf", "rockb.ttf"),
        "embebida": {"base": (0.62, 0.72), "ancha": (0.90, 0.72), "angosta": (0.40, 0.72)},
    },
    "bookman": {
        "nombre": "Bookman Old Style",
        "archivos": ("bookos.ttf", "bookosb.ttf"),
        "embebida": {"base": (0.72, 0.72), "ancha": (1.00, 0.72), "angosta": (0.42, 0.72)},
    },
}

# Margen de seguridad sobre la medición: exacta mide la fuente real que
# Word va a usar (cubre la negrita simulada); sustituta agrega holgura
# porque Word puede sustituir por otra más ancha; embebida es el último
# recurso sin medición real.
FACTOR_EXACTA = 1.10
FACTOR_SUSTITUTA = 1.18
FACTOR_EMBEBIDA = 1.12

# Márgenes de página (twips) y aire de recorte alrededor de la tinta.
MARGEN_SUP = 284
MARGEN_INF = 567
MARGEN_IZQ = 284
MARGEN_DER = 284
PIE_DIST = 227
# Aire que sobra en cada página para que las filas de una misma página
# apilen pese a los bordes, y el párrafo final entre en la última.
HOLGURA_PAGINA = 60
PAD_LATERAL_CM = 0.35
# Sonda empírica en Word (2024): con interlineado exacto igual al alto de
# fila, la tinta se ancla ~2,8 cm arriba del borde inferior y ESCAPA por
# arriba cuando la capital supera ~0,78 del alto de fila. 0.76 deja aire.
RAZON_ALTO_CENTRADO = 0.76

PT_POR_CM = 28.3465
TWIPS_POR_CM = 566.929

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


def _escapar(texto):
    from xml.sax.saxutils import escape

    return escape(texto)


def _campo_pie(instruccion, texto_previo):
    """Campo simple de Word (PAGE o NUMPAGES) con el estilo del pie."""
    rpr = (
        '<w:rPr><w:color w:val="' + COLOR_PIE + '"/>'
        + '<w:sz w:val="14"/><w:szCs w:val="14"/></w:rPr>'
    )
    return (
        '<w:fldSimple w:instr=" ' + instruccion + ' ">'
        + "<w:r>" + rpr + "<w:t>" + _escapar(texto_previo) + "</w:t></w:r>"
        + "</w:fldSimple>"
    )


def _pie_xml(texto):
    """Pie de página: identifica el cartel y numera las hojas."""
    rpr = (
        '<w:rPr><w:color w:val="' + COLOR_PIE + '"/>'
        + '<w:sz w:val="14"/><w:szCs w:val="14"/></w:rPr>'
    )
    return (
        CABECERA_XML
        + '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>'
        + "<w:r>" + rpr + '<w:t xml:space="preserve">'
        + _escapar(texto + " · Página ")
        + "</w:t></w:r>"
        + _campo_pie("PAGE", "1")
        + "<w:r>" + rpr + '<w:t xml:space="preserve"> de </w:t></w:r>'
        + _campo_pie("NUMPAGES", "1")
        + "</w:p></w:ftr>"
    )


def _letra_run_xml(caracter, sz_half, fuente_nombre, rellena):
    """Run de una letra: hueca (contorno, <w:outline/>) o rellena negra."""
    contorno = "" if rellena else "<w:outline/>"
    return (
        "<w:r><w:rPr>"
        + '<w:rFonts w:ascii="' + fuente_nombre + '" w:hAnsi="' + fuente_nombre + '" w:cs="' + fuente_nombre + '"/>'
        + "<w:b/><w:bCs/>"
        + contorno
        + '<w:color w:val="000000"/>'
        + '<w:sz w:val="' + str(sz_half) + '"/><w:szCs w:val="' + str(sz_half) + '"/>'
        + "</w:rPr>"
        + '<w:t xml:space="preserve">' + _escapar(caracter) + "</w:t></w:r>"
    )


def _celda_xml(caracter, sz_half, fuente_nombre, rellena, ancho, alto_fila):
    """Una celda de la grilla: letra centrada, o vacía si sobra lugar."""
    if caracter is None:
        parrafo = (
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>'
            + '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
        )
    else:
        parrafo = (
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="'
            + str(alto_fila)
            + '" w:lineRule="exact"/><w:jc w:val="center"/>'
            + '<w:rPr><w:sz w:val="' + str(sz_half) + '"/></w:rPr></w:pPr>'
            + _letra_run_xml(caracter, sz_half, fuente_nombre, rellena)
            + "</w:p>"
        )
    return (
        "<w:tc><w:tcPr>"
        + '<w:tcW w:w="' + str(ancho) + '" w:type="dxa"/>'
        + '<w:vAlign w:val="center"/>'
        + "</w:tcPr>"
        + parrafo
        + "</w:tc>"
    )


def _reparto_entero(total, partes):
    """Divide total en partes enteras, el resto en la última (determinista)."""
    if partes <= 0:
        raise ErrorGeneracion("Reparto con cantidad de partes no positiva.", 3)
    base = total // partes
    if base <= 0:
        raise ErrorGeneracion("Página demasiado chica para la grilla pedida.", 3)
    reparto = [base] * partes
    reparto[-1] = total - base * (partes - 1)
    return reparto


def _documento_xml(config):
    """Cuerpo: hojas × filas de columnas; cada fila exacta y cada hoja llena."""
    columnas = config["columnas"]
    filas = config["filas"]
    sz_half = config["sz_half"]
    fuente = config["fuente_nombre"]
    anchos = _reparto_entero(config["ancho_util"], columnas)
    patron_altos = _reparto_entero(config["alto_util"] - HOLGURA_PAGINA, filas)
    letras = config["letras"]

    filas_xml = []
    total_filas = config["hojas"] * filas
    for fila_global in range(total_filas):
        alto = patron_altos[fila_global % filas]
        celdas_xml = []
        for indice_col in range(columnas):
            posicion = fila_global * columnas + indice_col
            caracter = letras[posicion] if posicion < len(letras) else None
            celdas_xml.append(
                _celda_xml(caracter, sz_half, fuente, config["rellena"], anchos[indice_col], alto)
            )
        filas_xml.append(
            "<w:tr><w:trPr>"
            + '<w:trHeight w:val="' + str(alto) + '" w:hRule="exact"/>'
            + "<w:cantSplit/>"
            + "</w:trPr>"
            + "".join(celdas_xml)
            + "</w:tr>"
        )

    def _borde(nombre):
        return (
            '<w:' + nombre + ' w:val="dashed" w:sz="4" w:space="0" w:color="'
            + COLOR_GUIA + '"/>'
        )

    grid = "".join('<w:gridCol w:w="' + str(a) + '"/>' for a in anchos)
    tabla = (
        "<w:tbl><w:tblPr>"
        + '<w:tblW w:w="' + str(config["ancho_util"]) + '" w:type="dxa"/>'
        + '<w:tblLayout w:type="fixed"/>'
        + "<w:tblBorders>"
        + _borde("top") + _borde("left") + _borde("bottom") + _borde("right")
        + _borde("insideH") + _borde("insideV")
        + "</w:tblBorders>"
        + '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        + '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        + "</w:tblPr>"
        + "<w:tblGrid>" + grid + "</w:tblGrid>"
        + "".join(filas_xml)
        + "</w:tbl>"
        + '<w:p><w:pPr><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
    )

    if config["apaisada"]:
        pgsz = (
            '<w:pgSz w:w="' + str(config["pg_w"]) + '" w:h="' + str(config["pg_h"])
            + '" w:orient="landscape"/>'
        )
    else:
        pgsz = '<w:pgSz w:w="' + str(config["pg_w"]) + '" w:h="' + str(config["pg_h"]) + '"/>'
    sectpr = (
        "<w:sectPr>"
        + '<w:footerReference w:type="default" r:id="rId2"/>'
        + pgsz
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


def _partes_docx(config):
    """Las 6 partes del paquete, en orden fijo."""
    return (
        ("[Content_Types].xml", CONTENT_TYPES_XML),
        ("_rels/.rels", RELS_RAIZ_XML),
        ("word/document.xml", _documento_xml(config)),
        ("word/styles.xml", STYLES_XML),
        ("word/footer1.xml", _pie_xml(config["pie"])),
        ("word/_rels/document.xml.rels", RELS_DOCUMENTO_XML),
    )


def _docx_bytes(config):
    import io
    import zipfile

    buffer = io.BytesIO()
    try:
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as paquete:
            for nombre, contenido in _partes_docx(config):
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


# ---------------------------------------------------------------- métricas de letra

def _caracteres_de(texto):
    """Letras del cartel, sin espacios y sin repetidos, en orden de aparición."""
    caracteres = []
    for caracter in texto:
        if caracter.isspace():
            continue
        if caracter not in caracteres:
            caracteres.append(caracter)
    return caracteres


def _clase_tinta(caracter):
    """Clasificación anatómica del carácter para la tabla embebida."""
    if caracter in ("M", "W", "M", "W", "@", "%"):
        return "ancha"
    if caracter in ("I", "J", "1", "!", ".", ",", ":", ";", "'", "|"):
        return "angosta"
    return "base"


def _gdi_disponible():
    import os

    return os.name == "nt"


def _medir_con_gdi(caracteres, fuente_nombre):
    """Mide la tinta de cada carácter con GDI, en fracciones del cuerpo (em).

    Devuelve ({caracter: (ancho_em, alto_em)}, sustituida) o None si GDI
    no está disponible. Usa el mismo motor de texto que Word en esa
    máquina: si la fuente pedida no existe, mide la sustituta que
    realmente se va a ver (sustituida=True).

    Nota: pide el contorno con GGO_NATIVE en vez de GGO_METRICS porque
    llena el mismo GLYPHMETRICS y resultó más confiable; con buffer
    NULL solo consulta el tamaño y las métricas viajan igual.
    """
    if not _gdi_disponible():
        return None
    import ctypes
    from ctypes import wintypes

    FWCY = wintypes.WCHAR * 32
    em_px = 1000

    class LOGFONTW(ctypes.Structure):
        _fields_ = [
            ("lfHeight", wintypes.LONG),
            ("lfWidth", wintypes.LONG),
            ("lfEscapement", wintypes.LONG),
            ("lfOrientation", wintypes.LONG),
            ("lfWeight", wintypes.LONG),
            ("lfItalic", wintypes.BYTE),
            ("lfUnderline", wintypes.BYTE),
            ("lfStrikeOut", wintypes.BYTE),
            ("lfCharSet", wintypes.BYTE),
            ("lfOutPrecision", wintypes.BYTE),
            ("lfClipPrecision", wintypes.BYTE),
            ("lfQuality", wintypes.BYTE),
            ("lfPitchAndFamily", wintypes.BYTE),
            ("lfFaceName", FWCY),
        ]

    class POINTL(ctypes.Structure):
        _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

    class GLYPHMETRICS(ctypes.Structure):
        _fields_ = [
            ("gmBlackBoxX", wintypes.UINT),
            ("gmBlackBoxY", wintypes.UINT),
            ("gmptGlyphOrigin", POINTL),
            ("gmCellIncX", ctypes.c_short),
            ("gmCellIncY", ctypes.c_short),
        ]

    class MAT2(ctypes.Structure):
        # FIXED es un punto fijo 16.16 de 4 bytes; la matriz identidad
        # (1.0 en eM11 y eM22, 0.0 en el resto) equivale a estos DWORD.
        _fields_ = [
            ("eM11", wintypes.DWORD),
            ("eM12", wintypes.DWORD),
            ("eM21", wintypes.DWORD),
            ("eM22", wintypes.DWORD),
        ]

    mat2 = MAT2(65536, 0, 0, 65536)

    try:
        gdi32 = ctypes.windll.gdi32
        user32 = ctypes.windll.user32
        hdc = user32.GetDC(0)
        if not hdc:
            return None
        logfont = LOGFONTW()
        logfont.lfHeight = -em_px
        logfont.lfWeight = 700
        logfont.lfCharSet = 1
        logfont.lfOutPrecision = 4
        logfont.lfClipPrecision = 2
        logfont.lfQuality = 4
        logfont.lfPitchAndFamily = 34
        logfont.lfFaceName = fuente_nombre
        hfont = gdi32.CreateFontIndirectW(ctypes.byref(logfont))
        if not hfont:
            user32.ReleaseDC(0, hdc)
            return None
        viejo = gdi32.SelectObject(hdc, hfont)
        if not viejo:
            gdi32.DeleteObject(hfont)
            user32.ReleaseDC(0, hdc)
            return None
        mat2_listo = mat2
        GGO_NATIVE = 3
        GDI_ERROR = 0xFFFFFFFF
        mediciones = {}
        fallo = False
        for caracter in caracteres:
            metrica = GLYPHMETRICS()
            resultado = gdi32.GetGlyphOutlineW(
                hdc,
                wintypes.UINT(ord(caracter)),
                wintypes.UINT(GGO_NATIVE),
                ctypes.byref(metrica),
                wintypes.DWORD(0),
                None,
                ctypes.byref(mat2_listo),
            )
            if resultado == GDI_ERROR or metrica.gmBlackBoxX == 0 or metrica.gmBlackBoxY == 0:
                fallo = True
                break
            mediciones[caracter] = (
                metrica.gmBlackBoxX / float(em_px),
                metrica.gmBlackBoxY / float(em_px),
            )
        gdi32.SelectObject(hdc, viejo)
        gdi32.DeleteObject(hfont)
        user32.ReleaseDC(0, hdc)
        if fallo or len(mediciones) != len(caracteres):
            return None
        return mediciones
    except Exception:
        return None


def _medir_embebido(caracteres, preset_fuente):
    """Último recurso: fracciones de em embebidas, con margen extra."""
    tabla = {}
    for caracter in caracteres:
        tabla[caracter] = preset_fuente["embebida"][_clase_tinta(caracter)]
    return tabla


def _medir_tinta(caracteres, preset_fuente):
    """Tabla {caracter: (ancho_em, alto_em)} con el factor de seguridad incluido."""
    gdi = _medir_con_gdi(caracteres, preset_fuente["nombre"])
    if gdi is not None:
        mediciones = gdi
        instalada = _fuente_instalada(preset_fuente)
        factor = FACTOR_EXACTA if instalada else FACTOR_SUSTITUTA
        tabla = {}
        for caracter, (ancho_em, alto_em) in mediciones.items():
            tabla[caracter] = (ancho_em * factor, alto_em * factor)
        if instalada:
            origen = "GDI sobre la fuente real del equipo"
        else:
            origen = "GDI sobre fuente sustituta (la pedida no está instalada)"
        return tabla, origen
    tabla = {}
    for caracter in caracteres:
        ancho_em, alto_em = preset_fuente["embebida"][_clase_tinta(caracter)]
        tabla[caracter] = (ancho_em * FACTOR_EMBEBIDA, alto_em * FACTOR_EMBEBIDA)
    return (
        tabla,
        "métricas embebidas con margen extra (fuente no verificable)",
    )


# ---------------------------------------------------------------- grilla óptima

def _tamanio_de_grilla(caracteres, tinta, columnas, filas, ancho_util_pt, alto_util_pt):
    """Mayor tamaño de fuente (pt) que entra centrada en cada celda."""
    pad_lateral_pt = PAD_LATERAL_CM * PT_POR_CM
    celda_w = ancho_util_pt / float(columnas)
    celda_h = alto_util_pt / float(filas)
    tamanio = None
    for caracter in caracteres:
        ancho_em, alto_em = tinta[caracter]
        por_ancho = (celda_w - pad_lateral_pt) / ancho_em
        por_alto = RAZON_ALTO_CENTRADO * celda_h / alto_em
        candidato = min(por_ancho, por_alto)
        if tamanio is None or candidato < tamanio:
            tamanio = candidato
    return tamanio


def _elegir_grilla(caracteres, tinta, letras_por_hoja, ancho_util_pt, alto_util_pt):
    """Prueba todas las grillas válidas y elige la de letra más grande.

    Grilla válida: columnas×filas con columnas*filas >= N letras y sin
    desperdiciar una fila o columna entera. Empates: menos celdas
    vacías, luego más columnas (orden de lectura), luego menos filas.
    """
    n = letras_por_hoja
    candidatos = []
    for filas in range(1, n + 1):
        for columnas in range(1, n + 1):
            celdas = columnas * filas
            if celdas < n:
                continue
            if celdas - n >= min(columnas, filas):
                continue
            tamanio = _tamanio_de_grilla(
                caracteres, tinta, columnas, filas, ancho_util_pt, alto_util_pt
            )
            vacias = celdas - n
            candidatos.append((tamanio, -vacias, columnas, -filas, columnas, filas))
    if len(candidatos) == 0:
        raise ErrorGeneracion(
            "No hay ninguna grilla posible para " + str(n) + " letras por hoja.", 3
        )
    candidatos.sort(reverse=True)
    mejor = candidatos[0]
    if mejor[0] < TAMANO_MINIMO_PT:
        raise ErrorGeneracion(
            "El texto no entra: la letra más grande posible sería de menos de "
            + str(TAMANO_MINIMO_PT)
            + " pt. Reducí las letras por hoja, usá una hoja más grande o un texto más corto.",
            3,
        )
    return mejor[4], mejor[5], int(mejor[0])


# ---------------------------------------------------------------- configuración

def _normalizar_texto(bruto):
    """Mayúsculas, sin espacios en los extremos; valida el contenido."""
    texto = bruto.strip().upper()
    if len(texto) == 0:
        raise ErrorGeneracion("El texto del cartel está vacío.", 3)
    if len(texto) > TEXTO_MAXIMO:
        raise ErrorGeneracion(
            "Texto demasiado largo (" + str(len(texto)) + " caracteres, máximo "
            + str(TEXTO_MAXIMO) + ").",
            3,
        )
    for caracter in texto:
        if caracter.isspace():
            continue
        if ord(caracter) < 33 or ord(caracter) > 0x17F:
            raise ErrorGeneracion(
                "Carácter no imprimible en el cartel: " + repr(caracter), 3
            )
    return texto


def _pagina_de(pedido_hoja, pedido_ancho, pedido_alto, apaisada):
    """(ancho_cm, alto_cm) de la página según preset o medida libre.

    Con medida libre, ancho y alto son los de la hoja final; si las
    medidas contradicen la orientación pedida, el programa acomoda
    (apaisada nunca puede ser más alta que ancha).
    """
    if pedido_hoja == "medida":
        try:
            ancho = float(str(pedido_ancho).replace(",", "."))
            alto = float(str(pedido_alto).replace(",", "."))
        except (TypeError, ValueError):
            raise ErrorGeneracion(
                "Medida propia inválida: ancho y alto deben ser números en cm.", 3
            ) from None
        if min(ancho, alto) <= 5.0 or max(ancho, alto) > 200.0:
            raise ErrorGeneracion(
                "Medidas fuera de rango: se permiten entre 5 y 200 cm por lado.", 3
            )
    else:
        if pedido_hoja not in PRESETS_HOJA:
            raise ErrorGeneracion("Preset de hoja desconocido: " + str(pedido_hoja), 3)
        corto, largo = PRESETS_HOJA[pedido_hoja]
        ancho, alto = corto, largo
    if apaisada:
        if alto > ancho:
            ancho, alto = alto, ancho
    else:
        if ancho > alto:
            ancho, alto = alto, ancho
    return ancho, alto


def _fuente_instalada(preset_fuente):
    """True si el archivo de la fuente existe realmente en esta máquina.

    Exige el archivo, no solo la entrada del registro: una entrada
    huérfana haría medir la sustituta con factor de fuente real.
    """
    if not _gdi_disponible():
        return False
    import os
    import winreg

    carpetas = (
        os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts"),
        os.path.join(
            os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"
        ),
    )
    archivos = tuple(a.lower() for a in preset_fuente["archivos"])
    rutas = (
        "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts",
        "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Fonts",
    )
    for ruta in rutas:
        try:
            clave = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, ruta)
        except OSError:
            continue
        try:
            with clave:
                indice = 0
                while True:
                    try:
                        _, datos_valor, _ = winreg.EnumValue(clave, indice)
                    except OSError:
                        break
                    indice = indice + 1
                    nombre_archivo = os.path.basename(str(datos_valor)).lower()
                    if nombre_archivo not in archivos:
                        continue
                    if os.path.isabs(str(datos_valor)):
                        candidatos = (str(datos_valor),)
                    else:
                        candidatos = tuple(
                            os.path.join(carpeta, nombre_archivo) for carpeta in carpetas
                        )
                    if any(os.path.exists(candidato) for candidato in candidatos):
                        return True
        except OSError:
            continue
    return False


def _nombre_archivo(texto):
    import re
    import unicodedata

    descompuesto = unicodedata.normalize("NFKD", texto)
    plano = "".join(c for c in descompuesto if not unicodedata.combining(c))
    slug = re.sub(r"[^A-Z0-9]+", "-", plano).strip("-")
    if len(slug) == 0:
        slug = "cartel"
    return "banner-" + slug.lower() + "-letras-huecas.docx"


def _ruta_sin_pisar(ruta_destino):
    """Devuelve una ruta libre: si existe, agrega _V1, _V2, ... creciente."""
    import os

    if not os.path.exists(ruta_destino):
        return ruta_destino
    base, extension = os.path.splitext(ruta_destino)
    numero = 1
    while True:
        candidato = base + "_V" + str(numero) + extension
        if not os.path.exists(candidato):
            return candidato
        numero = numero + 1


def _carpeta_salida():
    import os
    import sys

    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _armar_config(texto, apaisada, hoja, ancho_cm, alto_cm, letras_por_hoja, clave_fuente, rellena, espacios):
    """Valida todo y devuelve el config completo listo para _docx_bytes."""
    import math

    preset = FUENTES_PRESET.get(clave_fuente)
    if preset is None:
        raise ErrorGeneracion("Fuente desconocida: " + str(clave_fuente), 3)
    if espacios not in ("quitar", "mantener"):
        raise ErrorGeneracion("Opción de espacios desconocida: " + str(espacios), 3)
    texto = _normalizar_texto(texto)
    if espacios == "mantener":
        letras = list(texto)
    else:
        letras = [c for c in texto if not c.isspace()]
    if len(letras) == 0:
        raise ErrorGeneracion("El texto no tiene ninguna letra para imprimir.", 3)
    if letras_por_hoja < 1 or letras_por_hoja > LETRAS_POR_HOJA_MAX:
        raise ErrorGeneracion(
            "Letras por hoja fuera de rango (1 a " + str(LETRAS_POR_HOJA_MAX) + ").", 3
        )
    ancho_cm_real, alto_cm_real = _pagina_de(hoja, ancho_cm, alto_cm, apaisada)
    pg_w = int(round(ancho_cm_real * TWIPS_POR_CM))
    pg_h = int(round(alto_cm_real * TWIPS_POR_CM))
    ancho_util = pg_w - MARGEN_IZQ - MARGEN_DER
    alto_util = pg_h - MARGEN_SUP - MARGEN_INF
    if ancho_util <= 0 or alto_util <= 0:
        raise ErrorGeneracion("La hoja es más chica que los márgenes.", 3)

    caracteres = _caracteres_de(texto)
    tinta, origen_metrica = _medir_tinta(caracteres, preset)
    columnas, filas, sz_pt = _elegir_grilla(
        caracteres,
        tinta,
        letras_por_hoja,
        ancho_util / TWIPS_POR_CM * PT_POR_CM,
        alto_util / TWIPS_POR_CM * PT_POR_CM,
    )

    hojas = int(math.ceil(len(letras) / float(letras_por_hoja)))
    return {
        "texto": texto,
        "letras": letras,
        "apaisada": apaisada,
        "pg_w": pg_w,
        "pg_h": pg_h,
        "ancho_util": ancho_util,
        "alto_util": alto_util,
        "columnas": columnas,
        "filas": filas,
        "sz_half": sz_pt * 2,
        "sz_pt": sz_pt,
        "fuente_nombre": preset["nombre"],
        "fuente_clave": clave_fuente,
        "origen_metrica": origen_metrica,
        "hojas": hojas,
        "letras_por_hoja": letras_por_hoja,
        "hoja_pedido": hoja,
        "rellena": bool(rellena),
        "espacios": espacios,
        "pie": texto + (" · letras huecas para recortar y colorear" if not rellena else " · letras para recortar"),
    }


def _verificar_xml(config):
    from defusedxml import ElementTree

    for nombre_parte, contenido in _partes_docx(config):
        try:
            ElementTree.fromstring(contenido.encode("utf-8"))
        except Exception as exc:
            raise ErrorGeneracion(
                "XML mal formado :: " + nombre_parte + ": " + repr(exc), 5
            ) from exc


def _generar(config, ruta_destino):
    corrida_a = _docx_bytes(config)
    corrida_b = _docx_bytes(config)
    if corrida_a != corrida_b:
        raise ErrorGeneracion(
            "Determinismo roto: dos corridas del mismo cartel difieren ("
            + str(len(corrida_a))
            + " vs "
            + str(len(corrida_b))
            + " bytes).",
            4,
        )
    _verificar_xml(config)
    try:
        with open(ruta_destino, "wb") as salida:
            salida.write(corrida_a)
    except OSError as exc:
        raise ErrorGeneracion(
            "No pude escribir " + ruta_destino + ": " + repr(exc), 6
        ) from exc
    return len(corrida_a)


# ---------------------------------------------------------------- consola

def _pedir(opcion_texto, defecto):
    """Pide un dato por consola; Enter toma el defecto, Ctrl+Z/Ctrl+D también."""
    try:
        respuesta = input(opcion_texto).strip()
    except (EOFError, KeyboardInterrupt):
        return defecto
    if len(respuesta) == 0:
        return defecto
    return respuesta


def _elegir_entre(opcion_texto, defecto, validas, error):
    respuesta = _pedir(opcion_texto, defecto)
    while respuesta not in validas:
        print("  " + error)
        respuesta = _pedir(opcion_texto, defecto)
    return respuesta


def _pedir_datos_interactivo():
    """Flujo de doble clic: pregunta de a uno, con valores por defecto."""
    print("Cartel de letras para recortar y colorear (Word 2010)")
    print("Enter acepta el valor entre corchetes.")
    print("")

    texto = _pedir("Texto del cartel [" + TEXTO_POR_DEFECTO + "]: ", TEXTO_POR_DEFECTO)

    orientacion = _elegir_entre(
        "Orientacion - [V]ertical o [A]paisada [A]: ",
        "a",
        ("v", "V", "a", "A"),
        "  Respondé V (vertical) o A (apaisada).",
    )
    apaisada = orientacion.lower() == "a"

    hoja = _elegir_entre(
        "Hoja: [1] A4  [2] A5  [3] Letter  [4] Medida propia [1]: ",
        "1",
        ("1", "2", "3", "4"),
        "  Elegí 1, 2, 3 o 4.",
    )
    hoja = {"1": "a4", "2": "a5", "3": "letter", "4": "medida"}[hoja]
    ancho_cm = None
    alto_cm = None
    if hoja == "medida":
        ancho_cm = _pedir("  Ancho en cm de la hoja [30.0]: ", "30.0")
        alto_cm = _pedir("  Alto en cm de la hoja [21.0]: ", "21.0")

    letras_raw = _elegir_entre(
        "Letras por hoja (1 a " + str(LETRAS_POR_HOJA_MAX) + ") [2]: ",
        "2",
        tuple(str(n) for n in range(1, LETRAS_POR_HOJA_MAX + 1)),
        "  Elegí un número entre 1 y " + str(LETRAS_POR_HOJA_MAX) + ".",
    )
    letras_por_hoja = int(letras_raw)

    fuente = _elegir_entre(
        "Fuente: [1] Arial Black  [2] Impact  [3] Arial  [4] Rockwell  [5] Bookman [1]: ",
        "1",
        ("1", "2", "3", "4", "5"),
        "  Elegí 1, 2, 3, 4 o 5.",
    )
    clave_fuente = {
        "1": "arial-black",
        "2": "impact",
        "3": "arial",
        "4": "rockwell",
        "5": "bookman",
    }[fuente]

    estilo = _elegir_entre(
        "Estilo: [H]ueca para colorear o [R]ellena negra [H]: ",
        "h",
        ("h", "H", "r", "R"),
        "  Respondé H (hueca) o R (rellena).",
    )
    rellena = estilo.lower() == "r"

    espacios = _elegir_entre(
        "Espacios: [Q]uitar (letras sueltas) o [M]antener (cartel continuo) [Q]: ",
        "q",
        ("q", "Q", "m", "M"),
        "  Respondé Q (quitar) o M (mantener).",
    )
    espacios = "mantener" if espacios.lower() == "m" else "quitar"

    return {
        "texto": texto,
        "apaisada": apaisada,
        "hoja": hoja,
        "ancho": ancho_cm,
        "alto": alto_cm,
        "letras": letras_por_hoja,
        "fuente": clave_fuente,
        "rellena": rellena,
        "espacios": espacios,
        "salida": None,
    }


def _resumen(config, ruta, bytes_generados):
    import os

    print("")
    print("OK " + os.path.basename(ruta) + " (" + str(bytes_generados) + " bytes)")
    estilo_texto = "rellena negra" if config["rellena"] else "hueca"
    detalle_espacios = " · espacios mantenidos" if config["espacios"] == "mantener" else ""
    print(
        "Texto: " + config["texto"]
        + " · " + str(len(config["letras"])) + " caracteres en "
        + str(config["hojas"]) + " hoja(s) · "
        + str(config["letras_por_hoja"]) + " por hoja · grilla "
        + str(config["columnas"]) + "x" + str(config["filas"])
        + detalle_espacios
    )
    for indice in range(config["hojas"]):
        porcion = config["letras"][indice * config["letras_por_hoja"]:(indice + 1) * config["letras_por_hoja"]]
        marcas = " ".join("·" if c.isspace() else c for c in porcion)
        print("  Hoja " + str(indice + 1) + ": " + marcas)
    print(
        "Letra: " + str(config["sz_pt"]) + " pt ("
        + str(round(config["sz_pt"] * 0.72 / PT_POR_CM, 1))
        + " cm de alto aprox.), " + config["fuente_nombre"] + " " + estilo_texto
    )
    print("Medición: " + config["origen_metrica"])
    print("Archivo: " + ruta)


def _construir_desde_args(argumentos):
    """Interpreta argv; con texto presente no pregunta nada."""
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        prog="banner-portable",
        description="Cartel de letras para recortar y colorear (DOCX Word 2010).",
    )
    parser.add_argument("texto", nargs="?", help="texto del cartel (por defecto " + TEXTO_POR_DEFECTO + ")")
    parser.add_argument("--vertical", action="store_true", help="hoja vertical")
    parser.add_argument("--apaisada", action="store_true", help="hoja apaisada (defecto)")
    parser.add_argument("--letras", type=int, default=2, help="letras por hoja, 1 a " + str(LETRAS_POR_HOJA_MAX) + " (defecto 2)")
    parser.add_argument("--hoja", choices=("a4", "a5", "letter", "medida"), default="a4", help="tamaño de hoja (defecto a4)")
    parser.add_argument("--ancho", help="ancho en cm de la hoja (con --hoja medida)")
    parser.add_argument("--alto", help="alto en cm de la hoja (con --hoja medida)")
    parser.add_argument("--fuente", choices=tuple(FUENTES_PRESET.keys()), default="arial-black", help="fuente de las letras (defecto arial-black)")
    parser.add_argument("--rellena", action="store_true", help="letra negra rellena en vez de hueca")
    parser.add_argument("--espacios", choices=("quitar", "mantener"), default="quitar", help="espacios entre palabras: quitar (letras sueltas, defecto) o mantener (cartel continuo)")
    parser.add_argument("-o", "--salida", help="ruta del DOCX de salida (si existe, agrega _V1, _V2...)")
    datos = parser.parse_args(argumentos)

    if datos.texto is not None:
        if datos.vertical and datos.apaisada:
            print("ERROR: --vertical y --apaisada son excluyentes.", file=sys.stderr)
            sys.exit(2)
        if datos.hoja == "medida" and (datos.ancho is None or datos.alto is None):
            print("ERROR: --hoja medida necesita --ancho y --alto en cm.", file=sys.stderr)
            sys.exit(2)
        return {
            "texto": datos.texto,
            "apaisada": not datos.vertical,
            "hoja": datos.hoja,
            "ancho": datos.ancho,
            "alto": datos.alto,
            "letras": datos.letras,
            "fuente": datos.fuente,
            "rellena": datos.rellena,
            "espacios": datos.espacios,
            "salida": datos.salida,
        }
    return _pedir_datos_interactivo()


def main(argumentos):
    import os
    import sys

    try:
        datos = _construir_desde_args(argumentos)
        config = _armar_config(
            datos["texto"], datos["apaisada"], datos["hoja"], datos["ancho"],
            datos["alto"], datos["letras"], datos["fuente"], datos["rellena"],
            datos["espacios"],
        )
        if datos["salida"]:
            ruta_pedida = datos["salida"]
        else:
            ruta_pedida = os.path.join(
                _carpeta_salida(), _nombre_archivo(config["texto"])
            )
        ruta_destino = _ruta_sin_pisar(ruta_pedida)
        bytes_generados = _generar(config, ruta_destino)
    except ErrorGeneracion as exc:
        print("ERROR: " + str(exc), file=sys.stderr)
        return exc.codigo
    _resumen(config, ruta_destino, bytes_generados)
    return 0


if __name__ == "__main__":
    import sys

    codigo_salida = main(sys.argv[1:])
    if getattr(sys, "frozen", False) and sys.stdin is not None and sys.stdin.isatty():
        import contextlib

        with contextlib.suppress(EOFError, KeyboardInterrupt):
            input("Enter para cerrar...")
    sys.exit(codigo_salida)
