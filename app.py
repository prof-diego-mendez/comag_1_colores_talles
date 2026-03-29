from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
import pandas as pd
import re
import os
import uuid
from werkzeug.utils import secure_filename

app = Flask(__name__)

# CORS configurado para producción
allowed_origins = os.environ.get('ALLOWED_ORIGINS', '*').split(',')
CORS(app, resources={r"/api/*": {"origins": allowed_origins}}, supports_credentials=True)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Diccionarios originales
mapeo_colores = {
    90: 'ALMENDRA', 15: 'AMARILLO', 534: 'AMATISTA', 810: 'AMBAR', 28: 'AMARILLO / NEGRO',
    49: 'ANIMAL PRINT', 964: 'AQUA - AGUAMARINA', 881: 'ARCILLA', 294: 'ARENA', 1007: 'ASTRO DUST',
    727: 'AVELLANA', 483: 'AZUL CLARO', 198: 'AZUL FRANCIA/ROYAL', 774: 'AZUL MELANGE', 40: 'AZUL MIX',
    78: 'AZUL NOCHE', 544: 'AZAFRAN', 667: 'AZUL-BLANCO', 973: 'AZUL/CRUDO', 775: 'AZUL GRIS',
    789: 'AZUL/NEGRO', 793: 'AZUL-ROSA', 309: 'AZUL-ROJO', 11: 'AZ/TURQUESA', 3: 'AZUL',
    580: 'AZUL ZAFIRO', 8: 'AZUL OSCURO', 611: 'AZUL PIEDRA', 997: 'AZUL PASTEL', 963: 'BAMBU',
    1001: 'B/C', 31: 'BLANCO', 785: 'BLANCO/FUCSIA', 150: 'BLANCO/GRIS', 14: 'BCO. PERLADO',
    156: 'BLANCO SATINADO', 37: 'BEIGE', 614: 'BEIGE MELANGE', 586: 'BEIGE/NARANJA', 670: 'BERENJENA',
    77: 'BERMELLON', 688: 'BLANCO-AMARILLO', 340: 'BLANCO-AQUA', 962: 'BLANCO - LILA', 850: 'BLANCO / AZUL',
    825: 'BLANCO / GRIS', 482: 'BLANCO / MARINO', 543: 'BLANCO/NEGRO', 805: 'BL/NE/VERDE', 866: 'BL/RO/BORDO',
    646: 'BLANCO / ROJO', 86: 'BLANCO/ROSA', 441: 'BLUSH', 669: 'BLANCO-VERDE', 140: 'BLANCO MATE',
    24: 'BORDO', 520: 'BORGOÑA', 536: 'BRANDY', 821: 'BRONCE', 972: 'BURDEO', 957: 'CAQUI',
    274: 'CACAO', 117: 'CAFE', 545: 'CAJU', 890: 'CALYPSO', 779: 'CAMELIA', 514: 'CAMEL',
    828: 'CAMUFLADO', 812: 'CANELA', 161: 'CAOBA', 553: 'CAPUCCINO', 112: 'CARMIN', 46: 'CARAMELO',
    606: 'CAREY', 826: 'CASTOR', 574: 'CASTAÑO', 829: 'CEBRA', 162: 'CEDRO', 58: 'CELESTE',
    6: 'CEMENTO', 766: 'CENIZA', 691: 'CEREZA', 186: 'CHICLE', 2: 'CHAMPAGNE', 960: 'CHERRY',
    68: 'CHOCOLATE', 717: 'CHUMBO', 22: 'CILANTRO', 280: 'CIRUELA', 981: 'CIELO', 265: 'CITRUS',
    125: 'COBRE', 976: 'COCOA', 51: 'COGNAC', 341: 'CORAZONES', 63: 'CORAL', 626: 'CREMA',
    999: 'CRISTAL', 860: 'CROCO/NATURAL', 439: 'CROCO/NE', 12: 'CRUDO', 163: 'CUADROS', 921: 'CUARZO',
    492: 'CURCUMA', 48: 'CURRY', 20: 'CYAN', 662: 'CEREZA', 628: 'Combo 1', 226: 'C10',
    282: 'C12', 650: 'C13', 768: 'C14', 678: 'C18', 299: 'Combo 2', 142: 'C21', 726: 'C22',
    83: 'C26', 837: 'C27', 311: 'C28', 876: 'Combo 3', 344: 'C4', 652: 'C5', 721: 'C6',
    281: 'C7', 651: 'C8', 392: 'C9', 731: 'DAMERO', 551: 'DEGRADEE', 882: 'DENIM', 524: 'DERBY',
    497: 'DORADO ROSA', 474: 'DORADO', 505: 'DURAZNO', 845: 'ESMERALDA', 87: 'ESTAMPADO', 205: 'ESTAMPADO NENA',
    204: 'ESTAMPADO VARON', 883: 'ESTAMPADO 1', 884: 'ESTAMPADO 2', 601: 'ETNICO', 1009: 'FANT 12', 1010: 'FANT 13',
    1011: 'FANT 14', 1012: 'FANT 15', 149: 'FANTASIA', 570: 'FANT 1', 833: 'FANT 2', 835: 'FANT 3',
    74: 'FANT 4', 336: 'FANT 5', 644: 'FANT 6', 164: 'FANT 7', 568: 'FANT 8', 1008: 'FANT 9',
    653: 'FANT 10', 569: 'FANT 11', 832: 'FANT 30', 283: 'FANT 32', 202: 'FLOREADO', 719: 'FRAMBUESA',
    54: 'FRANCIA', 421: 'FRUTILLA', 328: 'FROZEN', 73: 'FUCSIA', 290: 'GRIS HUMO', 838: 'GIRASOL',
    862: 'GLICINA', 965: 'GRIS/HIELO', 190: 'GRIS CLARO', 686: 'GRIS C/NUDE', 518: 'GRIS FROST', 203: 'GRIS JASPEADO',
    285: 'GRIS MICRO', 661: 'GRIS/NEGRO', 158: 'GRIS/GRAFITO OSCURO', 23: 'GRAFITO', 157: 'GRANITO', 979: 'GRANATE',
    682: 'GRIS-BLANCO', 209: 'GRIS/CELESTE', 767: 'GRIS FUCSIA', 934: 'GRIS/FRANCIA', 581: 'GRIS Y AZUL', 685: 'GRIS',
    588: 'GRIS MELANGE', 50: 'GRIS OSCURO', 585: 'GRIS PERLA', 608: 'GRIS-MIX', 799: 'GRIS/NEGRO/BLANCO', 151: 'GRIS OSCURO/ACERO',
    858: 'GRIS PLATA', 57: 'GRIS/ROJO', 189: 'GRIS-TUQUESA', 776: 'GRIS VERDE', 21: 'GUAYABA', 983: 'GUINDA',
    625: 'HABANO C/AVELL-BEIG', 617: 'HABANO C/ARABESCOS', 630: 'HABANO C/BEIGE', 618: 'HABANO C/GUARADA', 622: 'HABANOC/ROMBOS', 583: 'HABANO',
    41: 'HAVANNA', 67: 'HIELO', 772: 'HOJAS', 165: 'HORTENCIA', 994: 'HUESO', 297: 'HUMO',
    732: 'INDIGO', 604: 'ITALIA', 820: 'JADE', 527: 'JADE MARINO', 920: 'JADE PRETO', 552: 'JADE WHISKY',
    769: 'JASPEADO VE/GR', 975: 'JASPEADO AZ/GR', 16: 'JEAN', 210: 'JEAN/CANVAS', 493: 'JEAN/TABACO', 61: 'JENGIBRE',
    109: 'KAHKI/HABANO', 323: 'KITTY', 598: 'KIWI', 831: 'LACRE/LAQ', 961: 'LADRILLO', 533: 'LAUREL',
    887: 'LAVADO', 809: 'LAVANDA', 289: 'LECHUGA', 714: 'LEOPARDO GRIS', 318: 'LEOPARDO MARRON', 64: 'LILA',
    146: 'LIMON/LIMA', 127: 'LINO', 665: 'LILA-WHITE', 498: 'LLAMA', 490: 'LOVE', 479: 'LUNARES',
    42: 'MACADAMIA', 79: 'MACARRON', 711: 'MACAU', 331: 'MADERA', 18: 'MAGENTA', 599: 'MAIZ',
    827: 'MALBEC', 180: 'MALVA', 339: 'MANGO', 131: 'MANDARINA', 974: 'MANTECA', 549: 'VERDE MANZANA',
    52: 'MARRON', 718: 'MARACUYA', 546: 'MARFIL', 174: 'MARGARITA', 45: 'MARINO', 207: 'MARMOLADO',
    852: 'MARROCOS', 589: 'MARSALA', 30: 'MASCAVO', 635: 'MAUVE', 959: 'MANTECA/VIBORA', 532: 'MCE/MAMBO',
    199: 'MELON', 840: 'MEL', 181: 'MELANGE', 179: 'MELANGE OSCURO', 193: 'MENTA', 573: 'MERLOT',
    500: 'METAL', 521: 'MEZCLA ESCURO', 590: 'MEZCLA MEDIO', 327: 'MICKEY', 571: 'MIDNIGHT', 351: 'MIEL',
    195: 'MILITAR', 329: 'MINIONS', 324: 'MINNIE', 970: 'MIX/MENES/SABRO', 461: 'MOCA', 980: 'MOON',
    81: 'MORADO', 567: 'MORA', 629: 'MOSTAZA C/AZUL', 642: 'MOSTAZA', 200: 'MOUSE', 649: 'MULTICOLOR',
    889: 'MUSGO', 85: 'MOSTAZA/MAIZ', 168: 'CUADROS N', 34: 'NAPOLI HAVANNA', 36: 'NAPOLI SESAMO', 849: 'NAPA AMENDOA',
    952: 'NAPA', 684: 'NARANJA FLUO', 296: 'NARANJA PASTEL', 55: 'NARANJA', 7: 'NAPA ROJO', 32: 'NATURAL',
    516: 'NAVY', 664: 'NAVY-WHITE', 995: 'NARANJA CLARO', 723: 'NEGRO CHAROL', 643: 'NEGRO - NARANJA', 522: 'NEGRO ORO',
    971: 'NEGRO/AMARILLO', 139: 'NEGRO/BEIGE', 753: 'NEGRO Y BORDO', 33: 'NEGRO', 781: 'NEGRO C/AZUL', 542: 'NEGRO - BLANCO',
    208: 'NEGRO-SALMON', 494: 'NEGR0-VISON', 166: 'NE/GR/BLANCO', 822: 'NEGRO-CORAL', 638: 'NEGRO Y GRIS', 473: 'NEGRO/MARRON',
    978: 'NEGRO/PLATA', 515: 'NEGRO/ROJO/BLANCO', 477: 'NEGRO-ROSA', 676: 'NEGRO-TURQUESA', 926: 'NEGRO/WHISKY', 848: 'NAPA FLY BRANCA',
    471: 'NEGRO/BCO/CO', 496: 'NEG/GRIS/CELESTE', 778: 'NEGRO FUCSIA', 159: 'NEGRO MATE', 472: 'NEGRO/ROJO', 798: 'NIGHT',
    320: 'NOA', 38: 'NOBUCK MEL', 763: 'NOGAL', 44: 'NUDE', 639: 'NUEZ Y TOPO', 636: 'NUEZ',
    847: 'NAPA VERMELHA', 859: 'NVY', 459: 'NEW ROYAL', 460: 'NEW STONE', 169: 'CUADROS O', 555: 'OCEAN',
    554: 'OCRE', 851: 'OFF WHITE', 71: 'OLIVA', 454: 'OLIVA/NARANJA', 145: 'ORO', 114: 'OSCURO',
    641: 'OXFORD', 764: 'OXIDO', 855: 'PATINADO CLARO', 222: 'PANAL', 540: 'PATINADO', 648: 'PELTRE',
    335: 'PEPPA PIG', 279: 'PERA', 854: 'PEROLA', 5: 'PETROLEO', 286: 'PINO AZUL', 624: 'PIEL C/ETNICO',
    632: 'PIEL C/VERDE', 110: 'PIEDRA', 609: 'PIEDRA-MIX', 853: 'PIEL', 53: 'PINHAO', 984: 'PISTACHO',
    771: 'PIXEL PARTY', 582: 'PIZARRA', 332: 'PJ MASK', 819: 'PLATA - NARANJA', 65: 'PLATA/PLATEADO', 141: 'PLATA CROMO',
    728: 'PLOMO', 326: 'PLUMAS', 729: 'PO DE ARROZ', 762: 'PORCELANA', 615: 'PRALINE', 167: 'PRATA',
    25: 'PRETO', 325: 'PRINCESAS', 107: 'PUNTO', 519: 'PURPURA', 170: 'CUADROS Q', 843: 'RAINBOW',
    507: 'RANCHO', 720: 'RAYAS AZUL', 692: 'RAYAS VERDE', 201: 'RAYA 2', 842: 'RAYA 3', 844: 'RAYA 5',
    108: 'RAYADO', 634: 'RAYA 1', 841: 'RAYA 4', 284: 'REPTIL', 539: 'ROMBOS', 47: 'ROSA MAUVE',
    782: 'ROJO/AZUL', 668: 'ROSA-BLANCO', 878: 'ROBLE', 557: 'ROJO FLUO', 722: 'ROJO MELANGE', 898: 'ROJO PASTEL',
    13: 'ROJO', 687: 'ROJO LACRE', 814: 'ROJO OSC', 607: 'ROMANCE', 547: 'RON', 660: 'ROSA/NEGRO',
    788: 'ROJO/NEGRO', 647: 'ROSA CLARO', 558: 'ROSA FLUO', 683: 'ROSA VIEJO', 790: 'ROSA', 996: 'ROSA PASTEL',
    76: 'ROYAL', 66: 'SAFARI', 72: 'SALMON', 681: 'SALMON-PLATA', 287: 'SALVIA', 759: 'SAND',
    0: 'SIN COLOR', 115: 'SELVA', 824: 'SESAMO', 605: 'SHINE', 541: 'SIENA', 818: 'SILVER',
    212: 'SEGUN MUESTRA', 509: 'SUELA Y AZUL', 621: 'SUELA C/ARABESCOS', 631: 'SUELA C/DURAZNO', 619: 'SUELA C/GUARDA', 616: 'SUELA C/ZIG ZAG',
    640: 'SUELA Y HABANO', 637: 'SUELA', 993: 'SUMMER AZUL', 1: 'SUMMER', 172: 'SUMMER VERDE', 118: 'SURF',
    89: 'SURTIDO', 602: 'SURT.ESTAMPADO', 603: 'SURT.LISO', 633: 'TABACO C/CORAL', 620: 'TABACO C/GUARDA', 623: 'TABACO C/JASPEADO',
    508: 'TABACO', 846: 'TANNAT', 121: 'TAUPE', 292: 'TE VERDE', 823: 'TELA NEGRO/CRUDO', 504: 'TELA SUELA /CRUDO',
    645: 'TELA VIOLETA /CRUDO', 486: 'TELA VERDE LAUREL', 59: 'TELHA', 88: 'TERRACOTA', 560: 'TEXAS', 666: 'TEX/HABANO',
    488: 'TEX/TABACO', 484: 'THN', 75: 'TIERRA', 60: 'TIFFANY', 153: 'TIGRE', 113: 'TIZA',
    191: 'TOMATE', 506: 'TOPO', 360: 'TORTUGAS NINJA', 1000: 'TOSCANO', 503: 'TOSCANO', 62: 'TOSTADO',
    730: 'TOY STORY', 548: 'TRANSPARENTE', 386: 'TRICOLOR', 333: 'TROLLS', 817: 'TRUFA', 680: 'TURQUESA-CORAL',
    933: 'TULES', 35: 'TURIN PRETTO', 784: 'TURQ/FUCS', 56: 'TURQUESA', 322: 'TURRON', 584: 'UNICO',
    337: 'UNICORN', 111: 'UVA', 206: 'VAINILLA', 154: 'VERDE COLONIAL', 19: 'VERDE ESMERALDA', 587: 'VERDE PETRO',
    91: 'VERDE SECO', 591: 'VERDE LIMON', 9: 'VERDE MILITAR', 610: 'VERDE AGUA', 39: 'VERDE BOSQUE', 559: 'VERDE CLARO',
    43: 'VERDE FLUO', 80: 'VERDE GRANITO', 10: 'VERDE LIMA/LIMON', 613: 'VERDE LAUREL', 977: 'VERDE OLIVA', 816: 'VERDE OSCURO',
    556: 'VERDE PASTEL', 4: 'VERDE', 147: 'VERDE BENETTON', 787: 'VERDE INGLES', 958: 'VERDE SECO', 916: 'VERDE MUSGO',
    830: 'VIBORA', 398: 'VICHY', 780: 'VIOLETA FUCSIA', 143: 'VERDE INGLES', 70: 'VINO', 867: 'VIOLETA/BLANCO',
    17: 'VIOLETA', 155: 'VIOLETA PASTEL', 786: 'VILETA/ROJO', 760: 'VIOLETA/VERDE', 69: 'VISON', 144: 'VERDE MANZANA',
    293: 'VOSC', 998: 'VOYAGE', 82: 'WASHED BURGUNDY', 334: 'WHISKY', 517: 'WINE', 26: 'WOK HAVANNA', 369: 'ZIG ZAG'
}

mapeo_talles = {
    'XS': 0, 'S': 1, 'M': 2, 'L': 3, 'XL': 4, 'XXL': 5, 'XXXL': 6
}


def normalizar_texto(texto):
    """Convierte conectores (Y, &, CON) y símbolos (-, /, _) en espacios simples"""
    if pd.isna(texto):
        return ""
    texto_str = str(texto).upper()
    texto_str = re.sub(r'\bY\b|\bCON\b', ' ', texto_str)
    texto_str = re.sub(r'[-/_&]', ' ', texto_str)
    texto_str = re.sub(r'\s+', ' ', texto_str).strip()
    return texto_str


# Preparar el diccionario interno usando los nombres normalizados
mapeo_nombre_a_codigo = {}
for k, v in mapeo_colores.items():
    nombre_norm = normalizar_texto(v)
    mapeo_nombre_a_codigo[nombre_norm] = str(k).zfill(3)

# Ordenamos por longitud para dar prioridad a colores compuestos
nombres_colores_ordenados = sorted(mapeo_nombre_a_codigo.keys(), key=len, reverse=True)


def buscar_color_parcial(texto_color):
    if pd.isna(texto_color):
        return ""
    texto_limpio = normalizar_texto(texto_color)
    if texto_limpio in mapeo_nombre_a_codigo:
        return mapeo_nombre_a_codigo[texto_limpio]
    for nombre_conocido in nombres_colores_ordenados:
        if nombre_conocido in texto_limpio:
            return mapeo_nombre_a_codigo[nombre_conocido]
    return str(texto_color).strip().upper() + " (NO ENCONTRADO)"


def procesar_celda_talle(valor):
    if pd.isna(valor):
        return ""
    valor_str = str(valor).strip().upper()
    if not valor_str or valor_str == "NAN":
        return ""
    if valor_str in mapeo_talles:
        return str(mapeo_talles[valor_str]).zfill(3)
    else:
        try:
            numero_talle = int(float(valor_str))
            return str(numero_talle).zfill(3)
        except ValueError:
            return valor_str


def procesar_celda_barras(valor):
    if pd.isna(valor):
        return valor
    valor_str = str(valor).strip()
    if valor_str.endswith('.0') and valor_str[:-2].isdigit():
        return f"'{valor_str[:-2]}"
    if valor_str.isdigit():
        return f"'{valor_str}"
    return valor


def obtener_letra_columna(n):
    res = ""
    while n >= 0:
        res = chr(n % 26 + 65) + res
        n = n // 26 - 1
    return res


def procesar_excel(archivo_entrada, archivo_salida):
    """Procesa el archivo Excel y devuelve el path del archivo procesado"""
    df = pd.read_excel(archivo_entrada)

    columna_color = "color"
    columna_codigo_color = "codigo_color"
    columna_talle = "talle"
    columna_codigo_talle = "codigo_talle"
    columna_comag = "comag"
    columna_barras = "barras"

    # --- 1. PROCESAR LA COLUMNA COLOR ---
    if columna_color in df.columns:
        valores_codigo_color = df[columna_color].apply(buscar_color_parcial)
        indice_color = df.columns.get_loc(columna_color)
        df.insert(indice_color, columna_codigo_color, valores_codigo_color)
    else:
        print(f"Advertencia: No se encontró la columna '{columna_color}' en el Excel.")

    # --- 2. PROCESAR LA COLUMNA TALLE ---
    if columna_talle in df.columns:
        valores_codigo_talle = df[columna_talle].apply(procesar_celda_talle)
        indice_talle = df.columns.get_loc(columna_talle)
        df.insert(indice_talle, columna_codigo_talle, valores_codigo_talle)
    else:
        print(f"Advertencia: No se encontró la columna '{columna_talle}' en el Excel.")

    # --- 3. PROCESAR LA COLUMNA BARRAS ---
    if columna_barras in df.columns:
        df[columna_barras] = df[columna_barras].apply(procesar_celda_barras)

    # --- 4. CREAR LA COLUMNA CONCATENADA (FÓRMULA) AL PRINCIPIO ---
    if columna_comag in df.columns and columna_codigo_talle in df.columns and columna_codigo_color in df.columns:
        df[columna_comag] = df[columna_comag].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
        df.insert(0, 'Codigo_Completo', '')

        idx_comag = df.columns.get_loc(columna_comag)
        idx_talle_nuevo = df.columns.get_loc(columna_codigo_talle)
        idx_codigo_nuevo = df.columns.get_loc(columna_codigo_color)

        letra_comag = obtener_letra_columna(idx_comag)
        letra_talle = obtener_letra_columna(idx_talle_nuevo)
        letra_codigo = obtener_letra_columna(idx_codigo_nuevo)

        formulas = []
        for i in range(2, len(df) + 2):
            formulas.append(f'={letra_comag}{i}&{letra_talle}{i}&{letra_codigo}{i}')

        df['Codigo_Completo'] = formulas

    # Guardar el nuevo archivo
    df.to_excel(archivo_salida, index=False)
    return archivo_salida


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/procesar', methods=['POST'])
def procesar():
    if 'file' not in request.files:
        return jsonify({'error': 'No se encontró ningún archivo'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No se seleccionó ningún archivo'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Formato no válido. Solo se permiten archivos .xlsx'}), 400

    try:
        # Generar nombre único para el archivo
        unique_id = uuid.uuid4().hex
        filename = secure_filename(file.filename)
        nombre_base = filename.rsplit('.', 1)[0]
        archivo_entrada = os.path.join(UPLOAD_FOLDER, f"{unique_id}_{filename}")
        archivo_salida = os.path.join(UPLOAD_FOLDER, f"{unique_id}_{nombre_base}_actualizado.xlsx")

        # Guardar archivo temporal
        file.save(archivo_entrada)

        # Procesar el archivo
        resultado = procesar_excel(archivo_entrada, archivo_salida)

        # Limpiar archivo de entrada
        os.remove(archivo_entrada)

        return jsonify({
            'success': True,
            'message': 'Archivo procesado correctamente',
            'download_url': f'/api/descargar/{os.path.basename(resultado)}'
        })

    except Exception as e:
        return jsonify({'error': f'Error al procesar el archivo: {str(e)}'}), 500


@app.route('/api/descargar/<filename>')
def descargar(filename):
    return send_file(
        os.path.join(UPLOAD_FOLDER, filename),
        as_attachment=True
    )


@app.route('/api/limpiar', methods=['POST'])
def limpiar():
    """Limpia archivos temporales antiguos (más de 1 hora)"""
    import time
    ahora = time.time()
    eliminados = 0
    for archivo in os.listdir(UPLOAD_FOLDER):
        path = os.path.join(UPLOAD_FOLDER, archivo)
        if os.path.isfile(path) and (ahora - os.path.getmtime(path)) > 3600:
            os.remove(path)
            eliminados += 1
    return jsonify({'eliminados': eliminados})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    app.run(debug=debug_mode, host='0.0.0.0', port=port)
