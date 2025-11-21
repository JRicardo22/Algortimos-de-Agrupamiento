class XLSXtoCSV:

    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file


    def leer_archivo(self):
        with open(self.input_file, 'rb') as f:
            return f.read()


# ----------- Funciones para manejar ZIP ------------

def es_zip(self, data):
    # Los archivos ZIP comienzan con 0x50 0x4b 0x03 0x04
    return data[0:4] == b'\x50\x4b\x03\x04'


def parsear_cabeceras_zip(self, data):
    archivos = {}
    pos = 0
    while pos < len(data):
        if data[pos:pos + 4] != b'\x50\x4b\x03\x04':
            break
        header = data[pos:pos + 30]
        if len(header) < 30:
            break
        comp_method = int.from_bytes(header[8:10], 'little')
        comp_size = int.from_bytes(header[18:22], 'little')
        uncomp_size = int.from_bytes(header[22:26], 'little')
        fname_len = int.from_bytes(header[26:28], 'little')
        extra_len = int.from_bytes(header[28:30], 'little')
        name_start = pos + 30
        name_end = name_start + fname_len
        filename = data[name_start:name_end].decode('utf-8', errors='ignore')
        data_start = name_end + extra_len
        data_end = data_start + comp_size
        filedata = data[data_start:data_end]

        archivos[filename] = {
            'compression_method': comp_method,
            'compressed_data': filedata,
            'uncompressed_size': uncomp_size
        }
        pos = data_end
    return archivos


# ----------- BitStream para DEFLATE ------------

class BitStream:
    def __init__(self, data):
        self.data = data
        self.bitpos = 0
        self.bytepos = 0

    def leer_bit(self):
        if self.bytepos >= len(self.data):
            return None
        b = self.data[self.bytepos]
        bit = (b >> self.bitpos) & 1
        self.bitpos += 1
        if self.bitpos == 8:
            self.bitpos = 0
            self.bytepos += 1
        return bit

    def leer_bits(self, n):
        val = 0
        for i in range(n):
            bit = self.leer_bit()
            if bit is None:
                raise Exception("Fin inesperado de datos al leer bits")
            val |= (bit << i)
        return val


# ----------- Funciones Huffman ------------

def construir_tabla_huffman(self, longitudes):
    max_len = max(longitudes) if longitudes else 0
    bl_count = [0] * (max_len + 1)
    for l in longitudes:
        if l > 0:
            bl_count[l] += 1
    code = 0
    next_code = [0] * (max_len + 1)
    for bits in range(1, max_len + 1):
        code = (code + bl_count[bits - 1]) << 1
        next_code[bits] = code
    table = {}
    for n, length in enumerate(longitudes):
        if length != 0:
            code = next_code[length]
            next_code[length] += 1
            inv_code = 0
            for i in range(length):
                inv_code |= ((code >> i) & 1) << (length - 1 - i)
            table[(inv_code, length)] = n
    return table, max_len


def leer_codigo_huffman(self, bs, tabla, max_len):
    codigo = 0
    for longitud in range(1, max_len + 1):
        bit = bs.leer_bit()
        if bit is None:
            raise Exception("Fin inesperado en lectura huffman")
        codigo |= (bit << (longitud - 1))
        if (codigo, longitud) in tabla:
            return tabla[(codigo, longitud)]
    raise Exception("Codigo Huffman invalido")


# ----------- DEFLATE ------------

def descomprimir_deflate(self, data):
    bs = self.BitStream(data)
    salida = bytearray()
    final = False

    while not final:
        final = bs.leer_bit()
        tipo_bloque = bs.leer_bits(2)
        if tipo_bloque == 0:
            while bs.bitpos != 0:
                bs.leer_bit()
            len_bloque = bs.leer_bits(16)
            nlen_bloque = bs.leer_bits(16)
            if (len_bloque ^ 0xFFFF) != nlen_bloque:
                raise Exception("LEN y NLEN no coinciden")
            for _ in range(len_bloque):
                b = bs.leer_bits(8)
                salida.append(b)
        elif tipo_bloque == 1:
            salida.extend(self.descomprimir_huffman_fijo(bs))
        elif tipo_bloque == 2:
            salida.extend(self.descomprimir_huffman_dinamico(bs))
        else:
            raise Exception("Tipo de bloque desconocido")
    return bytes(salida)


def descomprimir_huffman_fijo(self, bs):
    salida = bytearray()
    long_lit = []
    for i in range(288):
        if i <= 143:
            long_lit.append(8)
        elif i <= 255:
            long_lit.append(9)
        elif i <= 279:
            long_lit.append(7)
        else:
            long_lit.append(8)
    tabla_lit, max_lit = self.construir_tabla_huffman(long_lit)
    long_dist = [5] * 32
    tabla_dist, max_dist = self.construir_tabla_huffman(long_dist)

    while True:
        simbolo = self.leer_codigo_huffman(bs, tabla_lit, max_lit)
        if simbolo == 256:
            break
        elif simbolo < 256:
            salida.append(simbolo)
        else:
            longitud = self.calcular_longitud(simbolo, bs)
            distancia = self.calcular_distancia(bs, tabla_dist, max_dist)
            for _ in range(longitud):
                salida.append(salida[-distancia])
    return salida


def descomprimir_huffman_dinamico(self, bs):
    salida = bytearray()
    HLIT = bs.leer_bits(5) + 257
    HDIST = bs.leer_bits(5) + 1
    HCLEN = bs.leer_bits(4) + 4
    orden_codigos_codigo = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]
    codigos_codigo_longitudes = [0] * 19
    for i in range(HCLEN):
        codigos_codigo_longitudes[orden_codigos_codigo[i]] = bs.leer_bits(3)
    tabla_codigo_codigo, max_codigo_codigo = self.construir_tabla_huffman(codigos_codigo_longitudes)

    def leer_codigo_codigo():
        return self.leer_codigo_huffman(bs, tabla_codigo_codigo, max_codigo_codigo)

    longitudes_ll = []
    while len(longitudes_ll) < HLIT + HDIST:
        c = leer_codigo_codigo()
        if c <= 15:
            longitudes_ll.append(c)
        elif c == 16:
            if not longitudes_ll: raise Exception("Error al repetir longitud")
            rep = bs.leer_bits(2) + 3
            longitudes_ll.extend([longitudes_ll[-1]] * rep)
        elif c == 17:
            rep = bs.leer_bits(3) + 3
            longitudes_ll.extend([0] * rep)
        elif c == 18:
            rep = bs.leer_bits(7) + 11
            longitudes_ll.extend([0] * rep)
        else:
            raise Exception("Codigo invalido en longitud huffman dinamico")

    long_lit = longitudes_ll[:HLIT]
    long_dist = longitudes_ll[HLIT:]
    tabla_lit, max_lit = self.construir_tabla_huffman(long_lit)
    tabla_dist, max_dist = self.construir_tabla_huffman(long_dist)

    while True:
        simbolo = self.leer_codigo_huffman(bs, tabla_lit, max_lit)
        if simbolo == 256:
            break
        elif simbolo < 256:
            salida.append(simbolo)
        else:
            longitud = self.calcular_longitud(simbolo, bs)
            distancia = self.calcular_distancia(bs, tabla_dist, max_dist)
            for _ in range(longitud):
                salida.append(salida[-distancia])
    return salida


# ----------- Cálculos longitud y distancia ------------

def calcular_longitud(self, simbolo, bs):
    long_tabla = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163,
                  195, 227, 258]
    extra_tabla = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0]
    index = simbolo - 257
    if index < 0 or index >= len(long_tabla): raise Exception("Simbolo longitud invalido")
    base = long_tabla[index]
    extra = extra_tabla[index]
    return base if extra == 0 else base + bs.leer_bits(extra)


def calcular_distancia(self, bs, tabla_dist, max_dist):
    dist_simbolo = self.leer_codigo_huffman(bs, tabla_dist, max_dist)
    dist_tabla = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073,
                  4097, 6145, 8193, 12289, 16385, 24577]
    dist_extra = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13]
    if dist_simbolo < 0 or dist_simbolo >= len(dist_tabla): raise Exception("Simbolo distancia invalido")
    base = dist_tabla[dist_simbolo]
    extra = dist_extra[dist_simbolo]
    return base if extra == 0 else base + bs.leer_bits(extra)


# ----------- Parseo XML minimalista ------------

def parsear_sharedStrings(self, data):
    strings = []
    texto = data.decode('utf-8', errors='ignore')
    pos = 0
    while True:
        start = texto.find("<t>", pos)
        if start == -1: break
        end = texto.find("</t>", start)
        if end == -1: break
        s = texto[start + 3:end]
        strings.append(s)
        pos = end + 4
    return strings


def parsear_sheet(self, data, shared_strings):
    texto = data.decode('utf-8', errors='ignore')
    filas = {}
    pos = 0
    while True:
        start_c = texto.find("<c ", pos)
        if start_c == -1: break
        end_c = texto.find(">", start_c)
        if end_c == -1: break
        c_tag = texto[start_c:end_c + 1]
        r_pos = c_tag.find("r=\"")
        if r_pos == -1:
            pos = end_c + 1
            continue
        r_start = r_pos + 3
        r_end = c_tag.find("\"", r_start)
        celda_ref = c_tag[r_start:r_end]
        t_pos = c_tag.find("t=\"")
        tipo = None
        if t_pos != -1:
            t_start = t_pos + 3
            t_end = c_tag.find("\"", t_start)
            tipo = c_tag[t_start:t_end]
        v_start = texto.find("<v>", end_c)
        if v_start == -1:
            pos = end_c + 1
            continue
        v_end = texto.find("</v>", v_start)
        if v_end == -1:
            pos = end_c + 1
            continue
        valor = texto[v_start + 3:v_end]
        if tipo == "s":
            try:
                valor = shared_strings[int(valor)]
            except:
                pass
        fila, col = 0, 0
        i_num = -1
        for i, ch in enumerate(celda_ref):
            if ch.isdigit():
                i_num = i
                break
        if i_num != -1:
            fila = int(celda_ref[i_num:])
            col_ref = celda_ref[:i_num]
        else:
            fila = 1
            col_ref = celda_ref
        col = 0
        for c in col_ref:
            col = col * 26 + (ord(c.upper()) - ord('A') + 1)
        if fila not in filas:
            filas[fila] = {}
        filas[fila][col] = valor
        pos = v_end + 4

    max_col = 0
    for f in filas.values():
        if f: max_col = max(max_col, max(f.keys()))

    lineas = []
    for f_num in sorted(filas.keys()):
        fila_data = filas[f_num]
        linea = [fila_data.get(c, "") for c in range(1, max_col + 1)]
        lineas.append(linea)
    return lineas


# ----------- Método para guardar CSV ------------

def guardar_csv(self, filas):
    with open(self.output_file, 'w', encoding='utf-8') as f:
        for fila in filas:
            linea = ','.join(['"{}"'.format(str(x).replace('"', '""')) for x in fila])
            f.write(linea + '\n')


# ----------- Método principal ------------

def procesar(self):
    datos = self.leer_archivo()
    if not self.es_zip(datos):
        raise Exception("Archivo no es ZIP (xlsx)")

    archivos = self.parsear_cabeceras_zip(datos)
    shared_strings = []
    if 'xl/sharedStrings.xml' in archivos:
        ss_info = archivos['xl/sharedStrings.xml']
        if ss_info['compression_method'] == 8:
            shared_strings_bytes = self.descomprimir_deflate(ss_info['compressed_data'])
        else:
            shared_strings_bytes = ss_info['compressed_data']
        shared_strings = self.parsear_sharedStrings(shared_strings_bytes)

    if 'xl/worksheets/sheet1.xml' not in archivos:
        raise Exception("No se encontró sheet1.xml en el XLSX")

    sheet_info = archivos['xl/worksheets/sheet1.xml']
    if sheet_info['compression_method'] == 8:
        sheet_bytes = self.descomprimir_deflate(sheet_info['compressed_data'])
    else:
        sheet_bytes = sheet_info['compressed_data']

    filas = self.parsear_sheet(sheet_bytes, shared_strings)
    self.guardar_csv(filas)


# ==============================================================================
#  PARTE 2: LÓGICA DE PROCESAMIENTO DE DATOS
#
# ==============================================================================

# --- Constantes para la limpieza y lectura de datos ---
POSIBLES_EXTENSIONES = ['', '.txt', '.csv', '.data', '.dat', '.xlsx']
DELIMITADORES_COMUNES = [',', ';', '\t']
VALORES_NULOS = ['na', 'null', 'none', 'n/a', 'vacio']

# --- Gestión de archivos temporales (sin librerías) ---
_contador_temporal_global = 0


def generar_nombre_temporal():
    """Genera un nombre de archivo temporal único usando un contador global."""
    global _contador_temporal_global
    _contador_temporal_global += 1
    return "temp_xlsx_convert_{}.csv".format(_contador_temporal_global)


def archivo_existe(nombre_archivo):
    """Verifica si un archivo existe sin usar la librería 'os'."""
    try:
        with open(nombre_archivo, 'r'):
            pass
        return True
    except FileNotFoundError:
        return False
    except IOError:  # Puede existir pero no ser legible
        return True


def eliminar_archivo_seguro(nombre_archivo):
    """Intenta eliminar un archivo. Falla silenciosamente si no es posible."""
    try:
        # Vaciar el archivo es una forma de "eliminar" su contenido.
        with open(nombre_archivo, 'w') as f:
            f.write('')
    except IOError:
        pass  # Ignorar errores si el archivo no se puede escribir/encontrar.


# --- Lógica de lectura y limpieza de datos ---

def _intentar_convertir_a_numero(valor_str):
    """
    Intenta convertir una cadena a un número (float).
    Prueba la conversión directa y luego reemplazando comas por puntos.
    """
    try:
        return float(valor_str)
    except (ValueError, TypeError):
        # Segundo intento: reemplazar comas por puntos
        if isinstance(valor_str, str):
            try:
                valor_corregido = valor_str.replace(',', '.')
                return float(valor_corregido)
            except (ValueError, TypeError):
                return valor_str  # Devuelve el original si falla
        return valor_str


def _procesar_valor_individual(valor_str):
    """Limpia, procesa y convierte un único valor de una celda."""
    if not isinstance(valor_str, str):
        return valor_str  # Ya es numérico o None

    valor_limpio = valor_str.strip()

    # Quitar comillas si envuelven el valor (común en CSV)
    if len(valor_limpio) >= 2 and valor_limpio.startswith('"') and valor_limpio.endswith('"'):
        valor_limpio = valor_limpio[1:-1]

    # Comprobar si es un valor nulo definido
    if valor_limpio == '' or valor_limpio.lower() in VALORES_NULOS:
        return None

    return _intentar_convertir_a_numero(valor_limpio)


def leer_datos(nombre_archivo_base):
    """
    Lee datos de un archivo, manejando varias extensiones, delimitadores y convirtiendo XLSX.
    """
    # 1. Encontrar el archivo real
    archivo_encontrado = None
    for ext in POSIBLES_EXTENSIONES:
        nombre_completo = nombre_archivo_base + ext
        if archivo_existe(nombre_completo):
            archivo_encontrado = nombre_completo
            break

    if not archivo_encontrado:
        raise FileNotFoundError(
            "No se pudo encontrar el archivo '{}' con ninguna de las extensiones: {}".format(nombre_archivo_base,
                                                                                             POSIBLES_EXTENSIONES))

    # 2. Convertir XLSX a un CSV temporal si es necesario
    archivo_temporal_csv = None
    if archivo_encontrado.lower().endswith('.xlsx'):
        archivo_temporal_csv = generar_nombre_temporal()
        try:
            print("Detectado archivo XLSX. Convirtiendo '{}' a CSV...".format(archivo_encontrado))
            convertidor = XLSXtoCSV(archivo_encontrado, archivo_temporal_csv)
            convertidor.procesar()
            print("Conversión exitosa. Procesando datos...")
            archivo_a_leer = archivo_temporal_csv
        except Exception as e:
            if archivo_temporal_csv:
                eliminar_archivo_seguro(archivo_temporal_csv)
            raise Exception("Error al convertir archivo XLSX: {}".format(str(e)))
    else:
        archivo_a_leer = archivo_encontrado

    # 3. Leer y procesar el archivo de texto (CSV, TXT, etc.)
    try:
        with open(archivo_a_leer, 'r', encoding='utf-8') as f:
            lineas = f.readlines()

        if not lineas:
            raise ValueError("El archivo está vacío.")

        # 4. Detección de encabezado (misma lógica original para mantener el resultado)
        primera_linea_valores = lineas[0].strip().replace(',', ' ').replace(';', ' ').replace('\t', ' ').split()
        es_encabezado = True
        for v in primera_linea_valores:
            try:
                float(v)
                es_encabezado = False  # Si al menos uno es número, no es encabezado
                break
            except ValueError:
                continue

        datos_crudos = lineas[1:] if es_encabezado else lineas

        # 5. Procesar cada fila y normalizar datos
        datos_procesados = []
        for linea in datos_crudos:
            linea = linea.strip()
            if not linea: continue

            # Detección de delimitador
            datos_fila = None
            for delim in DELIMITADORES_COMUNES:
                if delim in linea:
                    datos_fila = linea.split(delim)
                    break
            if datos_fila is None:
                datos_fila = linea.split()  # Espacio como último recurso

            fila_limpia = [_procesar_valor_individual(v) for v in datos_fila]
            datos_procesados.append(fila_limpia)

        if not datos_procesados:
            raise ValueError("No se encontraron datos válidos en el archivo.")

        # 6. Normalizar la longitud de las filas
        max_cols = 0
        for fila in datos_procesados:
            if len(fila) > max_cols:
                max_cols = len(fila)

        datos_normalizados = []
        for fila in datos_procesados:
            diferencia = max_cols - len(fila)
            if diferencia > 0:
                fila.extend([None] * diferencia)
            datos_normalizados.append(fila)

        return datos_normalizados

    finally:
        # 7. Limpiar archivo temporal si se creó
        if archivo_temporal_csv:
            eliminar_archivo_seguro(archivo_temporal_csv)


# --- Funciones de cálculo estadístico ---

def raiz_cuadrada_manual(numero):
    """Calcula la raíz cuadrada usando el método de Newton (requerido por no usar librerías)."""
    if numero < 0:
        raise ValueError("No se puede calcular la raíz cuadrada de un número negativo.")
    if numero == 0:
        return 0

    x = numero
    precision = 1e-10
    while True:
        raiz = 0.5 * (x + numero / x)
        if abs(raiz - x) < precision:
            break
        x = raiz
    return raiz


def calcular_media(columna):
    """Calcula la media de una lista, ignorando valores no numéricos."""
    numeros = [v for v in columna if isinstance(v, (int, float))]
    if not numeros: return None
    return sum(numeros) / len(numeros)


def calcular_desviacion_estandar(columna, media):
    """Calcula la desviación estándar poblacional."""
    numeros = [v for v in columna if isinstance(v, (int, float))]
    if not numeros or media is None or len(numeros) < 2:
        return 0.0 if len(numeros) == 1 else None

    suma_dif_cuadrado = sum((valor - media) ** 2 for valor in numeros)
    varianza = suma_dif_cuadrado / len(numeros)
    return raiz_cuadrada_manual(varianza)


def calcular_puntaje_z(datos):
    """Calcula los puntajes Z para cada columna numérica de la matriz de datos."""
    if not datos or not datos[0]:
        raise ValueError("Datos de entrada inválidos para el cálculo de puntaje Z.")

    num_filas = len(datos)
    num_cols = len(datos[0])

    # 1. Transponer datos para trabajar por columnas
    columnas = []
    for i_col in range(num_cols):
        columna = [datos[i_fila][i_col] for i_fila in range(num_filas)]
        columnas.append(columna)

    # 2. Calcular media y desviación estándar para cada columna
    medias = []
    desv_est = []
    for columna in columnas:
        media = calcular_media(columna)
        desviacion = calcular_desviacion_estandar(columna, media)
        medias.append(media)
        desv_est.append(desviacion)

    # 3. Calcular puntajes Z
    matriz_z = []
    for fila in datos:
        fila_z = []
        for i_col, valor in enumerate(fila):
            media_col = medias[i_col]
            desv_col = desv_est[i_col]

            if isinstance(valor, (int, float)) and media_col is not None and desv_col is not None and desv_col > 0:
                puntaje_z = (valor - media_col) / desv_col
                fila_z.append(puntaje_z)
            else:
                fila_z.append(valor)  # Mantener valores no numéricos o de columnas sin varianza
        matriz_z.append(fila_z)

    return matriz_z, medias, desv_est


# --- Presentación de resultados ---

def mostrar_resultados(datos_originales, matriz_z, medias, desv_est):
    """Muestra los resultados de forma clara y formateada, idéntica al original."""
    num_cols = len(datos_originales[0]) if datos_originales else 0
    num_filas = len(datos_originales)

    print("=" * 60)
    print("RESULTADOS DEL CÁLCULO DE PUNTAJE Z")
    print("=" * 60)

    print("\nDATOS ORIGINALES:")
    for i, fila in enumerate(datos_originales):
        fila_formateada = []
        for valor in fila:
            if valor is None:
                fila_formateada.append("    N/A ")
            elif isinstance(valor, (int, float)):
                fila_formateada.append("{:8.3f}".format(valor))
            else:
                fila_formateada.append("{:>8s}".format(str(valor)))
        print("Fila {:2d}: {}".format(i + 1, ' '.join(fila_formateada)))

    print("\nESTADÍSTICAS POR COLUMNA:")
    for i_col in range(num_cols):
        media = medias[i_col]
        desviacion = desv_est[i_col]

        # Conteo de tipos (más eficiente que en el original)
        num_numericos = 0
        num_no_numericos = 0
        for i_fila in range(num_filas):
            valor = datos_originales[i_fila][i_col]
            if isinstance(valor, (int, float)):
                num_numericos += 1
            else:
                num_no_numericos += 1

        if media is not None and desviacion is not None:
            print("Columna {:2d}: Media = {:10.6f}, Desv. Estándar = {:10.6f}".format(i_col + 1, media, desviacion))
            print("            ({} valores numéricos, {} valores de texto)".format(num_numericos, num_no_numericos))
        else:
            print("Columna {:2d}: Sin datos numéricos para calcular estadísticas".format(i_col + 1))
            print("            ({} valores numéricos, {} valores de texto)".format(num_numericos, num_no_numericos))

    print("\nMATRIZ DE PUNTAJES Z:")
    for i, fila in enumerate(matriz_z):
        fila_formateada = []
        for valor in fila:
            if valor is None:
                fila_formateada.append("    N/A ")
            elif isinstance(valor, (int, float)):
                # Lógica de formato idéntica a la original
                if abs(valor) < 10:
                    fila_formateada.append("{:8.3f}".format(valor))
                else:
                    fila_formateada.append("{:8.1f}".format(valor))
            else:
                fila_formateada.append("{:>8s}".format(str(valor)))
        print("Fila {:2d}: {}".format(i + 1, ' '.join(fila_formateada)))

    print("\n" + "=" * 60)


# --- Función principal ---

def main():
    """Función principal que orquesta la ejecución del script."""
    print("Calculadora de Puntaje Z")
    print("=" * 40)
    print("\nNota: Puedes escribir solo 'iris' y el programa buscará automáticamente")
    print("archivos con extensiones comunes (.txt, .csv, .data, .dat, .xlsx)")

    nombre_archivo = input("\nIngrese el nombre del archivo de datos: ").strip()

    try:
        print("\nBuscando archivo '{}'...".format(nombre_archivo))
        datos = leer_datos(nombre_archivo)
        print("Se cargaron exitosamente {} filas con {} columnas.".format(len(datos), len(datos[0])))

        print("\nCalculando puntajes Z...")
        matriz_z, medias, desv_est = calcular_puntaje_z(datos)

        print("\nMostrando resultados...")
        mostrar_resultados(datos, matriz_z, medias, desv_est)

    except FileNotFoundError as e:
        print("Error: {}".format(e))
    except ValueError as e:
        print("Error de Datos: {}".format(e))
    except Exception as e:
        print("Error Inesperado: {}".format(e))


if __name__ == "__main__":
    main()