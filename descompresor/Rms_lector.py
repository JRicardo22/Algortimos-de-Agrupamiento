class LectorXLSXCSVError(Exception):
    """Excepción base para errores del lector XLSX/CSV."""
    pass


class LectorXLSXCSV:
    """
    Lector minimalista de archivos .xlsx (Office Open XML) y .csv sin librerías externas.
    - Para .csv: devuelve una lista de filas (lista de listas).
    - Para .xlsx: descomprime el ZIP, ubica sharedStrings y sheet1.xml, y parsea las celdas.
    """

    def __init__(self, archivo_entrada, archivo_salida=None):
        if not isinstance(archivo_entrada, str) or not archivo_entrada.strip():
            raise LectorXLSXCSVError("La ruta del archivo de entrada es inválida o está vacía.")
        self.archivo_entrada = archivo_entrada
        self.archivo_salida = archivo_salida

    # ----------- Utilidades de lectura básica ------------

    def leer_archivo(self):
        """Lee el archivo binario completo."""
        try:
            with open(self.archivo_entrada, "rb") as f:
                datos = f.read()
            if not datos:
                raise LectorXLSXCSVError("El archivo está vacío.")
            return datos
        except FileNotFoundError as e:
            raise LectorXLSXCSVError(f"No se encontró el archivo: {self.archivo_entrada}") from e
        except OSError as e:
            raise LectorXLSXCSVError(f"Error al leer el archivo: {e}") from e

    # ----------- Funciones para manejar ZIP ------------

    @staticmethod
    def es_zip(datos):
        """Verifica si los primeros bytes corresponden a un archivo ZIP."""
        if not datos or len(datos) < 4:
            return False
        # ZIP files start with 0x50 0x4b 0x03 0x04
        return datos[0:4] == b"\x50\x4b\x03\x04"

    @staticmethod
    def parsear_cabeceras_zip(datos):
        """
        Parseo simple de cabeceras locales de ZIP (no soporta todos los casos posibles,
        pero es suficiente para archivos XLSX típicos).
        """
        archivos = {}
        pos = 0
        largo = len(datos)

        while pos + 30 <= largo:
            # Firma de cabecera local
            if datos[pos:pos + 4] != b"\x50\x4b\x03\x04":
                break

            header = datos[pos:pos + 30]
            if len(header) < 30:
                break

            metodo_compresion = int.from_bytes(header[8:10], "little")
            tam_comp = int.from_bytes(header[18:22], "little")
            tam_descomp = int.from_bytes(header[22:26], "little")
            longitud_nombre = int.from_bytes(header[26:28], "little")
            longitud_extra = int.from_bytes(header[28:30], "little")

            nombre_inicio = pos + 30
            nombre_fin = nombre_inicio + longitud_nombre
            if nombre_fin > largo:
                # Cabecera corrupta
                raise LectorXLSXCSVError(
                    "Cabecera ZIP corrupta (nombre de archivo fuera de rango)."
                )

            try:
                nombre_archivo = datos[nombre_inicio:nombre_fin].decode(
                    "utf-8", errors="ignore"
                )
            except Exception as e:
                raise LectorXLSXCSVError(
                    f"Error al decodificar nombre de archivo en ZIP: {e}"
                ) from e

            datos_inicio = nombre_fin + longitud_extra
            datos_fin = datos_inicio + tam_comp
            if datos_fin > largo:
                raise LectorXLSXCSVError(
                    f"Cabecera ZIP corrupta para '{nombre_archivo}' "
                    f"(datos comprimidos fuera de rango)."
                )

            datos_comprimidos = datos[datos_inicio:datos_fin]

            archivos[nombre_archivo] = {
                "metodo_compresion": metodo_compresion,
                "datos_comprimidos": datos_comprimidos,
                "tam_descomprimido": tam_descomp,
            }

            pos = datos_fin

        if not archivos:
            raise LectorXLSXCSVError("No se encontraron entradas válidas en el ZIP.")
        return archivos

    # ----------- BitStream para DEFLATE ------------

    class FlujoBits:
        def __init__(self, datos):
            self.datos = datos
            self.bitpos = 0
            self.bytepos = 0

        def leer_bit(self):
            if self.bytepos >= len(self.datos):
                raise LectorXLSXCSVError("Fin inesperado de datos al leer un bit.")
            b = self.datos[self.bytepos]
            bit = (b >> self.bitpos) & 1
            self.bitpos += 1
            if self.bitpos == 8:
                self.bitpos = 0
                self.bytepos += 1
            return bit

        def leer_bits(self, n):
            if n <= 0:
                return 0
            valor = 0
            for i in range(n):
                bit = self.leer_bit()
                valor |= (bit << i)
            return valor

    # ----------- Funciones Huffman ------------

    @staticmethod
    def construir_tabla_huffman(longitudes):
        if not longitudes:
            raise LectorXLSXCSVError("Lista de longitudes Huffman vacía.")

        max_long = max(longitudes)
        if max_long == 0:
            raise LectorXLSXCSVError("Todas las longitudes Huffman son cero.")

        conteo_long = [0] * (max_long + 1)
        for l in longitudes:
            if l < 0:
                raise LectorXLSXCSVError("Longitud Huffman negativa detectada.")
            if l > 0:
                conteo_long[l] += 1

        codigo = 0
        siguiente_codigo = [0] * (max_long + 1)
        for bits in range(1, max_long + 1):
            codigo = (codigo + conteo_long[bits - 1]) << 1
            siguiente_codigo[bits] = codigo

        tabla = {}
        for indice, longitud in enumerate(longitudes):
            if longitud != 0:
                codigo = siguiente_codigo[longitud]
                siguiente_codigo[longitud] += 1

                # Invertir bits para lectura LSB-first
                codigo_invertido = 0
                for i in range(longitud):
                    codigo_invertido |= ((codigo >> i) & 1) << (longitud - 1 - i)

                tabla[(codigo_invertido, longitud)] = indice

        return tabla, max_long

    @staticmethod
    def leer_codigo_huffman(flujo_bits, tabla, max_longitud):
        codigo = 0
        for longitud in range(1, max_longitud + 1):
            bit = flujo_bits.leer_bit()
            codigo |= (bit << (longitud - 1))
            clave = (codigo, longitud)
            if clave in tabla:
                return tabla[clave]
        raise LectorXLSXCSVError("Código Huffman inválido o no encontrado.")

    # ----------- DEFLATE ------------

    def descomprimir_deflate(self, datos):
        flujo = self.FlujoBits(datos)
        salida = bytearray()
        bloque_final = False

        while not bloque_final:
            bloque_final = bool(flujo.leer_bit())
            tipo_bloque = flujo.leer_bits(2)

            if tipo_bloque == 0:
                # Bloque sin comprimir
                while flujo.bitpos != 0:
                    flujo.leer_bit()  # Alinear a byte
                len_bloque = flujo.leer_bits(16)
                nlen_bloque = flujo.leer_bits(16)
                if (len_bloque ^ 0xFFFF) != nlen_bloque:
                    raise LectorXLSXCSVError(
                        "LEN y NLEN no coinciden en bloque sin comprimir."
                    )
                for _ in range(len_bloque):
                    b = flujo.leer_bits(8)
                    salida.append(b)

            elif tipo_bloque == 1:
                salida.extend(self._descomprimir_huffman_fijo(flujo))

            elif tipo_bloque == 2:
                salida.extend(self._descomprimir_huffman_dinamico(flujo))

            else:
                raise LectorXLSXCSVError(f"Tipo de bloque DEFLATE desconocido: {tipo_bloque}")

        return bytes(salida)

    def _descomprimir_huffman_fijo(self, flujo):
        salida = bytearray()

        longitudes_literal = []
        for i in range(288):
            if i <= 143:
                longitudes_literal.append(8)
            elif i <= 255:
                longitudes_literal.append(9)
            elif i <= 279:
                longitudes_literal.append(7)
            else:
                longitudes_literal.append(8)
        tabla_literal, max_literal = self.construir_tabla_huffman(longitudes_literal)

        longitudes_dist = [5] * 32
        tabla_dist, max_dist = self.construir_tabla_huffman(longitudes_dist)

        while True:
            simbolo = self.leer_codigo_huffman(flujo, tabla_literal, max_literal)
            if simbolo == 256:  # fin de bloque
                break
            elif simbolo < 256:
                salida.append(simbolo)
            else:
                longitud = self._calcular_longitud(simbolo, flujo)
                distancia = self._calcular_distancia(flujo, tabla_dist, max_dist)
                if distancia > len(salida):
                    raise LectorXLSXCSVError(
                        "Distancia de copia mayor que datos disponibles en salida "
                        "(bloque fijo)."
                    )
                for _ in range(longitud):
                    salida.append(salida[-distancia])

        return salida

    def _descomprimir_huffman_dinamico(self, flujo):
        salida = bytearray()

        HLIT = flujo.leer_bits(5) + 257
        HDIST = flujo.leer_bits(5) + 1
        HCLEN = flujo.leer_bits(4) + 4

        if HLIT > 286 or HDIST > 32:
            raise LectorXLSXCSVError("Valores HLIT/HDIST inválidos en bloque dinámico.")

        orden_codigos_codigo = [
            16, 17, 18, 0, 8, 7, 9, 6, 10,
            5, 11, 4, 12, 3, 13, 2, 14, 1, 15,
        ]
        longitudes_codigos_codigo = [0] * 19
        for i in range(HCLEN):
            longitudes_codigos_codigo[orden_codigos_codigo[i]] = flujo.leer_bits(3)

        tabla_codigos_codigo, max_codigos_codigo = self.construir_tabla_huffman(
            longitudes_codigos_codigo
        )

        def leer_codigo_codigo():
            return self.leer_codigo_huffman(flujo, tabla_codigos_codigo, max_codigos_codigo)

        longitudes_ll = []
        objetivo = HLIT + HDIST

        while len(longitudes_ll) < objetivo:
            c = leer_codigo_codigo()
            if c <= 15:
                longitudes_ll.append(c)
            elif c == 16:
                if not longitudes_ll:
                    raise LectorXLSXCSVError(
                        "Repetición de longitud sin valor previo (código 16)."
                    )
                rep = flujo.leer_bits(2) + 3
                longitudes_ll.extend([longitudes_ll[-1]] * rep)
            elif c == 17:
                rep = flujo.leer_bits(3) + 3
                longitudes_ll.extend([0] * rep)
            elif c == 18:
                rep = flujo.leer_bits(7) + 11
                longitudes_ll.extend([0] * rep)
            else:
                raise LectorXLSXCSVError("Código inválido en longitudes Huffman dinámico.")

            if len(longitudes_ll) > objetivo + 32:  # margen de seguridad
                raise LectorXLSXCSVError(
                    "Se generaron demasiadas longitudes Huffman dinámicas."
                )

        longitudes_literal = longitudes_ll[:HLIT]
        longitudes_dist = longitudes_ll[HLIT:]

        tabla_literal, max_literal = self.construir_tabla_huffman(longitudes_literal)
        tabla_dist, max_dist = self.construir_tabla_huffman(longitudes_dist)

        while True:
            simbolo = self.leer_codigo_huffman(flujo, tabla_literal, max_literal)
            if simbolo == 256:
                break
            elif simbolo < 256:
                salida.append(simbolo)
            else:
                longitud = self._calcular_longitud(simbolo, flujo)
                distancia = self._calcular_distancia(flujo, tabla_dist, max_dist)
                if distancia > len(salida):
                    raise LectorXLSXCSVError(
                        "Distancia de copia mayor que datos disponibles en salida "
                        "(bloque dinámico)."
                    )
                for _ in range(longitud):
                    salida.append(salida[-distancia])

        return salida

    # ----------- Cálculos longitud y distancia ------------

    @staticmethod
    def _calcular_longitud(simbolo, flujo):
        tabla_longitudes = [
            3, 4, 5, 6, 7, 8, 9, 10,
            11, 13, 15, 17, 19, 23, 27, 31,
            35, 43, 51, 59, 67, 83, 99, 115,
            131, 163, 195, 227, 258,
        ]
        tabla_extra = [
            0, 0, 0, 0, 0, 0, 0, 0,
            1, 1, 1, 1, 2, 2, 2, 2,
            3, 3, 3, 3, 4, 4, 4, 4,
            5, 5, 5, 5, 0,
        ]

        indice = simbolo - 257
        if indice < 0 or indice >= len(tabla_longitudes):
            raise LectorXLSXCSVError(f"Símbolo de longitud inválido: {simbolo}")

        base = tabla_longitudes[indice]
        extra = tabla_extra[indice]
        if extra == 0:
            return base
        return base + flujo.leer_bits(extra)

    @staticmethod
    def _calcular_distancia(flujo, tabla_dist, max_dist):
        simbolo_dist = LectorXLSXCSV.leer_codigo_huffman(flujo, tabla_dist, max_dist)

        tabla_distancias = [
            1, 2, 3, 4, 5, 7, 9, 13,
            17, 25, 33, 49, 65, 97, 129, 193,
            257, 385, 513, 769, 1025, 1537, 2049, 3073,
            4097, 6145, 8193, 12289, 16385, 24577,
        ]
        tabla_extra = [
            0, 0, 0, 0, 1, 1, 2, 2,
            3, 3, 4, 4, 5, 5, 6, 6,
            7, 7, 8, 8, 9, 9, 10, 10,
            11, 11, 12, 12, 13, 13,
        ]

        if simbolo_dist < 0 or simbolo_dist >= len(tabla_distancias):
            raise LectorXLSXCSVError(f"Símbolo de distancia inválido: {simbolo_dist}")

        base = tabla_distancias[simbolo_dist]
        extra = tabla_extra[simbolo_dist]
        if extra == 0:
            return base
        return base + flujo.leer_bits(extra)

    # ----------- Parseo XML minimalista para sharedStrings ------------

    @staticmethod
    def parsear_shared_strings(datos):
        """Extrae la lista de cadenas de sharedStrings.xml de forma muy simple."""
        try:
            texto = datos.decode("utf-8", errors="ignore")
        except Exception as e:
            raise LectorXLSXCSVError(f"Error al decodificar sharedStrings.xml: {e}") from e

        cadenas = []
        pos = 0
        while True:
            inicio = texto.find("<t", pos)
            if inicio == -1:
                break

            # Saltar posibles atributos: <t xml:space="preserve">
            inicio_contenido = texto.find(">", inicio)
            if inicio_contenido == -1:
                break

            fin = texto.find("</t>", inicio_contenido)
            if fin == -1:
                break

            contenido = texto[inicio_contenido + 1:fin]
            cadenas.append(contenido)
            pos = fin + 4

        return cadenas

    # ----------- Parseo XML minimalista para sheet1 ------------

    def parsear_sheet(self, datos, shared_strings):
        try:
            texto = datos.decode("utf-8", errors="ignore")
        except Exception as e:
            raise LectorXLSXCSVError(f"Error al decodificar sheet1.xml: {e}") from e

        filas = {}
        pos = 0
        largo = len(texto)

        while True:
            inicio_celda = texto.find("<c ", pos)
            if inicio_celda == -1:
                break
            fin_celda = texto.find(">", inicio_celda)
            if fin_celda == -1:
                break

            celda = texto[inicio_celda:fin_celda + 1]

            # Referencia, por ejemplo: r="A1"
            ref_inicio = celda.find('r="')
            if ref_inicio == -1:
                pos = fin_celda + 1
                continue
            ref_fin = celda.find('"', ref_inicio + 3)
            if ref_fin == -1:
                pos = fin_celda + 1
                continue
            referencia = celda[ref_inicio + 3:ref_fin]

            # Tipo de celda, ej: t="s" (shared string)
            tipo = None
            tipo_inicio = celda.find('t="')
            if tipo_inicio != -1:
                tipo_fin = celda.find('"', tipo_inicio + 3)
                if tipo_fin != -1:
                    tipo = celda[tipo_inicio + 3:tipo_fin]

            # Valor de la celda: <v>...</v>
            inicio_valor = texto.find("<v>", fin_celda)
            if inicio_valor == -1 or inicio_valor >= largo:
                pos = fin_celda + 1
                continue
            fin_valor = texto.find("</v>", inicio_valor)
            if fin_valor == -1:
                pos = fin_celda + 1
                continue

            valor_raw = texto[inicio_valor + 3:fin_valor]

            # Convertir referencia: letras -> fila, dígitos -> columna
            parte_letras = "".join(ch for ch in referencia if ch.isalpha())
            parte_numeros = "".join(ch for ch in referencia if ch.isdigit())

            if not parte_numeros:
                pos = fin_valor + 4
                continue

            try:
                fila_num = self._letra_a_numero(parte_letras)
                col_num = int(parte_numeros)
            except ValueError:
                pos = fin_valor + 4
                continue

            if fila_num <= 0 or col_num <= 0:
                pos = fin_valor + 4
                continue

            valor_final = valor_raw
            if tipo == "s":  # shared string
                try:
                    indice_ss = int(valor_raw)
                    valor_final = (
                        shared_strings[indice_ss] if 0 <= indice_ss < len(shared_strings) else ""
                    )
                except (ValueError, IndexError):
                    valor_final = ""

            if fila_num not in filas:
                filas[fila_num] = {}
            filas[fila_num][col_num] = valor_final

            pos = fin_valor + 4

        if not filas:
            # Hoja vacía: devolvemos []
            return []

        max_fila = max(filas.keys())
        max_col = 0
        for _, celdas in filas.items():
            if celdas:
                max_col = max(max_col, max(celdas.keys()))

        resultado = []
        for i in range(1, max_fila + 1):
            fila_actual = []
            for j in range(1, max_col + 1):
                fila_actual.append(filas.get(i, {}).get(j, ""))
            resultado.append(fila_actual)

        return resultado

    @staticmethod
    def _letra_a_numero(letras):
        """Convierte letras de columna Excel (A,B,...,AA,AB,...) a número 1-based."""
        if not letras:
            return 0
        letras = letras.upper()
        num = 0
        for c in letras:
            if "A" <= c <= "Z":
                num = num * 26 + (ord(c) - ord("A") + 1)
        return num

    # ----------- Procesamiento principal ------------

    def procesar(self):
        """
        Procesa el archivo de entrada:
        - Si es .csv: lo lee como texto.
        - Si es .xlsx: lo trata como ZIP + XML.
        Devuelve una lista de filas (lista de listas).
        """
        nombre = self.archivo_entrada.lower()

        if nombre.endswith(".csv"):
            return self._procesar_csv()

        # --- Procesamiento para XLSX ---
        datos = self.leer_archivo()
        if not self.es_zip(datos):
            raise LectorXLSXCSVError("El archivo no es un ZIP válido (no parece ser un XLSX).")

        archivos_zip = self.parsear_cabeceras_zip(datos)

        # sharedStrings.xml (opcional)
        if "xl/sharedStrings.xml" in archivos_zip:
            info_ss = archivos_zip["xl/sharedStrings.xml"]
            if info_ss["metodo_compresion"] == 8:  # DEFLATE
                bytes_ss = self.descomprimir_deflate(info_ss["datos_comprimidos"])
            else:
                bytes_ss = info_ss["datos_comprimidos"]
            shared_strings = self.parsear_shared_strings(bytes_ss)
        else:
            shared_strings = []

        # sheet1.xml (obligatorio para este lector)
        if "xl/worksheets/sheet1.xml" not in archivos_zip:
            raise LectorXLSXCSVError("No se encontró 'xl/worksheets/sheet1.xml' en el XLSX.")

        info_sheet = archivos_zip["xl/worksheets/sheet1.xml"]
        if info_sheet["metodo_compresion"] == 8:
            bytes_sheet = self.descomprimir_deflate(info_sheet["datos_comprimidos"])
        else:
            bytes_sheet = info_sheet["datos_comprimidos"]

        filas = self.parsear_sheet(bytes_sheet, shared_strings)
        return filas

    # ----------- Procesamiento CSV ------------

    def _procesar_csv(self):
        """Lee un CSV sencillo y lo devuelve como lista de filas (transpuesta, como en XLSX)."""
        try:
            with open(self.archivo_entrada, "r", encoding="utf-8") as f:
                contenido = f.read()
        except UnicodeDecodeError:
            # Intento de respaldo con latin-1
            with open(self.archivo_entrada, "r", encoding="latin-1", errors="ignore") as f:
                contenido = f.read()
        except FileNotFoundError as e:
            raise LectorXLSXCSVError(f"No se encontró el archivo CSV: {self.archivo_entrada}") from e
        except OSError as e:
            raise LectorXLSXCSVError(f"Error al leer el archivo CSV: {e}") from e

        if not contenido.strip():
            return []

        # Detección muy simple del separador
        separador = ","
        primera_linea = contenido.splitlines()[0]
        if primera_linea.count(";") > primera_linea.count(","):
            separador = ";"

        filas = []
        for linea in contenido.splitlines():
            linea = linea.strip()
            if not linea:
                continue
            partes = [col.strip().strip('"') for col in linea.split(separador)]
            filas.append(partes)

        if not filas:
            return []

        num_columnas = max(len(f) for f in filas)
        # Normalizamos el ancho de todas las filas
        filas_normalizadas = [
            fila + [""] * (num_columnas - len(fila))
            for fila in filas
        ]

        # Transponer (igual que tu versión original)
        filas_transpuestas = []
        for j in range(num_columnas):
            fila = []
            for i in range(len(filas_normalizadas)):
                fila.append(filas_normalizadas[i][j])
            filas_transpuestas.append(fila)

        return filas_transpuestas

    # ----------- Guardar como CSV ------------

    def guardar_csv(self, filas):
        """Guarda una lista de filas (lista de listas) como archivo CSV."""
        if not self.archivo_salida:
            raise LectorXLSXCSVError("No se especificó archivo de salida para CSV.")
        try:
            with open(self.archivo_salida, "w", encoding="utf-8", newline="") as f:
                for fila in filas:
                    # Escapar comillas dobles dentro de campos
                    campos = ['"{}"'.format(str(campo).replace('"', '""')) for campo in fila]
                    linea = ",".join(campos)
                    f.write(linea + "\n")
        except OSError as e:
            raise LectorXLSXCSVError(
                f"Error al escribir el archivo CSV de salida: {e}"
            ) from e
