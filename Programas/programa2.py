from zipfile import ZipFile
from xml.etree import ElementTree

FALTANTES = {"", "na", "nan", "null", "none"}  # se usa .strip().lower()
SEPARADORES_POSIBLES = [",", ";", "\t", "|", " "]

# ---------------- Entrada segura y pequeños helpers ----------------

def seguro_input(prompt, default=None, validar=None):
    while True:
        try:
            txt = input(prompt)
        except Exception:
            return default
        if validar is None or validar(txt):
            return txt
        print("Entrada inválida, intenta de nuevo.")

def to_lower(s):
    try:
        return s.lower()
    except Exception:
        return s

def es_faltante(celda):
    if celda is None:
        return True
    return to_lower(celda.strip()) in FALTANTES

def intentar_float(txt):
    try:
        return float(txt)
    except Exception:
        return None

def suma(valores):
    t = 0.0
    for v in valores:
        t += v
    return t

def maximo(valores):
    if not valores:
        return 0
    m = valores[0]
    for v in valores[1:]:
        if v > m:
            m = v
    return m

def minimo(valores):
    if not valores:
        return 0
    m = valores[0]
    for v in valores[1:]:
        if v < m:
            m = v
    return m

# ---------------- Detección tipo de archivo y lectura base ----------------

def leer_ruta_y_tipo():
    """
    Pide la ruta y detecta si es XLSX (intenta abrir como ZIP con [Content_Types].xml)
    o texto (abre como UTF-8). Devuelve (ruta, tipo, payload) donde payload es:
      - para xlsx: None (se vuelve a abrir dentro del lector xlsx)
      - para texto: lista de líneas no vacías
    """
    while True:
        ruta = seguro_input("Ruta del archivo (.xlsx o texto): ", default="")
        if not ruta:
            print("No se proporcionó ruta.")
            continue
        # ¿XLSX?
        try:
            with ZipFile(ruta, "r") as z:
                if "[Content_Types].xml" in z.namelist():
                    return ruta, "xlsx", None
        except Exception:
            pass
        # ¿Texto?
        try:
            with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
                lineas = [ln.rstrip("\n\r") for ln in f if ln.strip() != ""]
            if not lineas:
                print("El archivo está vacío o no tiene líneas útiles.")
                continue
            return ruta, "texto", lineas
        except Exception:
            print("No se pudo abrir el archivo como XLSX ni como texto. Verifica la ruta.")

# ---------------- Lectura de TEXTO (CSV/TSV/pipe/espacios) ----------------

def detectar_separador(lineas_muestra):
    mejor_sep = ","
    mejor_cols = 1
    mejor_consistencia = -1.0
    for sep in SEPARADORES_POSIBLES:
        conteos = []
        for ln in lineas_muestra:
            cols = partir_linea(ln, sep)
            conteos.append(len(cols))
        if not conteos:
            continue
        promedio = suma(conteos) / len(conteos)
        var = suma([(c - promedio) * (c - promedio) for c in conteos]) / len(conteos)
        consistencia = -var
        max_cols = maximo(conteos)
        if (max_cols > mejor_cols) or (max_cols == mejor_cols and consistencia > mejor_consistencia):
            mejor_cols = max_cols
            mejor_consistencia = consistencia
            mejor_sep = sep
    return mejor_sep

def pedir_separador(defecto):
    nombre = "tabulador" if defecto == "\t" else ("espacio" if defecto == " " else defecto)
    print("\nSeparador sugerido:", repr(nombre))
    resp = seguro_input("Forzar separador? (Enter=aceptar, ',', ';', '|', '\\t', ' '): ", default="")
    if not resp:
        return defecto
    if to_lower(resp) == "\\t":
        return "\t"
    if resp in {",", ";", "|", " "}:
        return resp
    print("Entrada no válida; usando el sugerido.")
    return defecto

def partir_linea(linea, sep):
    if sep == " ":
        partes = linea.strip().split()
    else:
        partes = linea.split(sep)
    limpias = []
    for p in partes:
        limpias.append(p.strip())
    return limpias

def lineas_a_tabla(lineas, sep):
    filas = []
    for l in lineas:
        if l.strip() == "":
            continue
        filas.append(partir_linea(l, sep))
    return filas

# ---------------- Lectura de XLSX (sin librerías externas) ----------------

def leer_xlsx_a_matriz(ruta, sheet_index=None):
    """
    Convierte un .xlsx a una matriz de strings (lista de filas).
    - sheet_index: 1..N (si None, pregunta si hay varias hojas).
    - Resuelve sharedStrings; otros valores van tal cual del nodo <v>.
    - No mapea estilos/fechas (se deja el número crudo).
    """
    with ZipFile(ruta, "r") as z:
        names = z.namelist()

        # sharedStrings
        shared = []
        if "xl/sharedStrings.xml" in names:
            try:
                sst = ElementTree.fromstring(z.read("xl/sharedStrings.xml"))
                for si in sst.iter():
                    if si.tag.endswith("si"):
                        txt = ""
                        for tnode in si.iter():
                            if tnode.tag.endswith("t") and tnode.text is not None:
                                txt += tnode.text
                        shared.append(txt)
            except Exception:
                shared = []

        # workbook y hojas
        sheets = []
        try:
            wb = ElementTree.fromstring(z.read("xl/workbook.xml"))
            for node in wb.iter():
                if node.tag.endswith("sheet"):
                    nm = node.attrib.get("name", "")
                    sheets.append(nm)
        except Exception:
            pass

        total = len([n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")])
        if total == 0:
            raise ValueError("XLSX sin hojas legibles.")

        if sheet_index is None:
            if total > 1 and sheets:
                print("\nHojas detectadas (1..{}):".format(total))
                for idx, nm in enumerate(sheets, start=1):
                    print("  {}: {}".format(idx, nm))
                ans = seguro_input("Elige hoja (Enter = 1): ", default="1")
                try:
                    sel = int(ans)
                except Exception:
                    sel = 1
            else:
                sel = 1
        else:
            sel = int(sheet_index)

        if sel < 1 or sel > total:
            sel = 1

        sheet_path = "xl/worksheets/sheet{}.xml".format(sel)
        if sheet_path not in names:
            cand = [n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
            cand.sort()
            sheet_path = cand[0]

        xml = ElementTree.fromstring(z.read(sheet_path))
        filas_dict = {}
        max_col_index = 0

        for c in xml.iter():
            if not c.tag.endswith("c"):
                continue
            r = c.attrib.get("r", "")
            t = c.attrib.get("t", "")
            v_node = None
            for sub in c.iter():
                if sub.tag.endswith("v"):
                    v_node = sub
                    break
            if not r:
                continue
            fila_idx, col_idx = ref_a_indices(r)
            if col_idx > max_col_index:
                max_col_index = col_idx

            val = ""
            if v_node is not None and v_node.text is not None:
                raw = v_node.text
                if t == "s":
                    try:
                        si = int(raw)
                        if 0 <= si < len(shared):
                            val = shared[si]
                        else:
                            val = raw
                    except Exception:
                        val = raw
                elif t == "b":
                    val = "1" if raw.strip() == "1" else "0"
                else:
                    val = raw
            else:
                val = ""

            if fila_idx not in filas_dict:
                filas_dict[fila_idx] = {}
            filas_dict[fila_idx][col_idx] = val

        if not filas_dict:
            return []

        max_row = max(filas_dict.keys())
        matriz = []
        for r_i in range(1, max_row + 1):
            fila_vals = []
            row_dict = filas_dict.get(r_i, {})
            if not row_dict:
                fila_vals = ["" for _ in range(max_col_index)]
            else:
                for c_j in range(1, max_col_index + 1):
                    fila_vals.append(row_dict.get(c_j, ""))
            matriz.append(fila_vals)

        return matriz

def ref_a_indices(ref):
    letras = ""
    numeros = ""
    for ch in ref:
        if ("A" <= ch <= "Z") or ("a" <= ch <= "z"):
            letras += ch
        elif "0" <= ch <= "9":
            numeros += ch
    col = letras_a_indice(letras)
    row = int(numeros) if numeros else 1
    return row, col

def letras_a_indice(letters):
    letters = letters.upper()
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

# ---------------- Normalización de tabla ----------------

def normalizar_ancho(filas):
    freq = {}
    for fila in filas:
        n = len(fila)
        freq[n] = freq.get(n, 0) + 1
    ncols_obj = None
    mejor = -1
    for k, v in freq.items():
        if v > mejor:
            mejor = v
            ncols_obj = k
    norm = []
    for fila in filas:
        if len(fila) < ncols_obj:
            fila = fila + [""] * (ncols_obj - len(fila))
        elif len(fila) > ncols_obj:
            fila = fila[:ncols_obj]
        norm.append(fila)
    return norm, ncols_obj

# ---------------- Selección de columnas ----------------

def mostrar_encabezado(encabezado):
    print("\nColumnas disponibles:")
    for i, nombre in enumerate(encabezado, start=1):
        print("  {}: {}".format(i, nombre))

def parsear_columna_usuario(txt, encabezado):
    if not txt:
        return None
    es_entero = True
    for ch in txt:
        if ch < "0" or ch > "9":
            es_entero = False
            break
    if es_entero:
        try:
            idx = int(txt)
            if 1 <= idx <= len(encabezado):
                return idx - 1
        except Exception:
            return None
    # nombre exacto
    for i, nm in enumerate(encabezado):
        if nm == txt:
            return i
    # nombre insensible a mayúsculas
    low = to_lower(txt)
    for i, nm in enumerate(encabezado):
        if to_lower(nm) == low:
            return i
    return None

# ---------------- Distancia euclidiana ----------------

def distancia_euclidiana_col(datos, col_a, col_b):
    suma_sq = 0.0
    usados = 0
    ignorados = 0
    for fila in datos:
        va = fila[col_a]; vb = fila[col_b]
        if es_faltante(va) or es_faltante(vb):
            ignorados += 1
            continue
        fa = intentar_float(va); fb = intentar_float(vb)
        if fa is None or fb is None:
            ignorados += 1
            continue
        dif = fa - fb
        suma_sq += dif * dif
        usados += 1
    if usados == 0:
        return 0.0, 0, ignorados
    # sqrt por método babilónico para evitar importar math
    x = suma_sq
    if x == 0.0:
        return 0.0, usados, ignorados
    g = x
    i = 0
    while i < 10:
        g = 0.5 * (g + x / g)
        i += 1
    return g, usados, ignorados

def formato_float(x):
    try:
        return "{:.6f}".format(float(x))
    except Exception:
        return "0.000000"

# ---------------- Flujo principal ----------------

def main():
    print("=== Distancia Euclidiana entre Columnas ===")
    ruta, tipo, payload = leer_ruta_y_tipo()

    if tipo == "xlsx":
        try:
            matriz = leer_xlsx_a_matriz(ruta, sheet_index=None)
        except Exception:
            print("No se pudo leer el XLSX. Verifica que no esté corrupto.")
            return
        if not matriz:
            print("El XLSX no aportó datos.")
            return

        ans = seguro_input("¿La primera fila es encabezado? [s/n] (Enter = 's'): ", default="s")
        tiene_encabezado = (str(ans).strip().lower() != "n")

        filas_norm, ncols = normalizar_ancho(matriz)
        if tiene_encabezado:
            encabezado = filas_norm[0]
            datos = filas_norm[1:]
        else:
            encabezado = []
            i = 0
            while i < ncols:
                encabezado.append("col_" + str(i + 1))
                i += 1
            datos = filas_norm

    else:
        lineas = payload
        muestra = lineas[:50]
        sep_detectado = detectar_separador(muestra)
        sep = pedir_separador(sep_detectado)
        ans = seguro_input("¿La primera fila es encabezado? [s/n] (Enter = 's'): ", default="s")
        tiene_encabezado = (str(ans).strip().lower() != "n")
        filas = lineas_a_tabla(lineas, sep)
        filas_norm, ncols = normalizar_ancho(filas)
        if tiene_encabezado:
            encabezado = filas_norm[0]
            datos = filas_norm[1:]
        else:
            encabezado = []
            i = 0
            while i < ncols:
                encabezado.append("col_" + str(i + 1))
                i += 1
            datos = filas_norm

    if not datos or ncols < 2:
        print("Datos insuficientes o menos de 2 columnas.")
        return

    mostrar_encabezado(encabezado)
    colA_txt = seguro_input("\nElige columna A (nombre o índice 1..{}): ".format(len(encabezado)), default="")
    colB_txt = seguro_input("Elige columna B (nombre o índice 1..{}): ".format(len(encabezado)), default="")

    idxA = parsear_columna_usuario(colA_txt, encabezado)
    idxB = parsear_columna_usuario(colB_txt, encabezado)
    if idxA is None or idxB is None or idxA == idxB:
        print("Selección de columnas inválida (revisa nombres/índices y que sean distintas).")
        return

    dist, usados, ignorados = distancia_euclidiana_col(datos, idxA, idxB)

    print("\n================= RESULTADO =================")
    print("Columna A:", encabezado[idxA], " (índice:", idxA + 1, ")")
    print("Columna B:", encabezado[idxB], " (índice:", idxB + 1, ")")
    print("Filas usadas (válidas):", usados)
    print("Filas ignoradas:", ignorados)
    print("Distancia euclidiana:", formato_float(dist))
    print("============================================\n")
    print("Listo.")

if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("Ocurrió un error no previsto. El programa finaliza de forma segura.")
