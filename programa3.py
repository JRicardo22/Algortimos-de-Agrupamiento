from zipfile import ZipFile
from xml.etree import ElementTree


class XLSXtoCSV:
    def __init__(self, input_file, output_file=None):
        self.input_file = input_file
        self.output_file = output_file

    # --- Utilidades ZIP mínimas (nos apoyamos en ZipFile de la stdlib) ---
    def procesar(self):
        # CSV simple (coma). Para otros delimitadores, lo leemos como texto abajo.
        if self.input_file.lower().endswith(".csv"):
            with open(self.input_file, 'r', encoding='utf-8', errors='ignore') as f:
                contenido = f.read()
            lineas = [ln for ln in contenido.splitlines() if ln.strip() != ""]
            filas = [ [c.strip().strip('"') for c in ln.split(",")] for ln in lineas ]
            return filas

        # XLSX: usamos ZipFile (robusto para tamaños/flags)
        with ZipFile(self.input_file, "r") as z:
            names = z.namelist()

            # sharedStrings (opcional)
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

            # workbook para listar hojas
            sheets = []
            try:
                wb = ElementTree.fromstring(z.read("xl/workbook.xml"))
                for node in wb.iter():
                    if node.tag.endswith("sheet"):
                        nm = node.attrib.get("name", "")
                        sheets.append(nm)
            except Exception:
                pass

            # preferimos sheet1.xml; si no está, tomamos la primera
            cand = [n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
            if not cand:
                raise Exception("XLSX sin hojas legibles.")
            cand.sort()
            sheet_path = "xl/worksheets/sheet1.xml" if "xl/worksheets/sheet1.xml" in names else cand[0]

            xml = ElementTree.fromstring(z.read(sheet_path))

            # Parseo de celdas con mapeo correcto row/col
            filas_dict = {}
            max_col_index = 0
            for c in xml.iter():
                if not c.tag.endswith("c"):
                    continue
                r = c.attrib.get("r", "")   # ej: "C5"
                t = c.attrib.get("t", "")   # tipo: s (shared), b (bool), ... (num/texto crudo)
                v_node = None
                for sub in c.iter():
                    if sub.tag.endswith("v"):
                        v_node = sub
                        break
                if not r:
                    continue

                # === CORRECCIÓN CLAVE ===
                letras = ""
                numeros = ""
                for ch in r:
                    if ("A" <= ch <= "Z") or ("a" <= ch <= "z"):
                        letras += ch
                    elif "0" <= ch <= "9":
                        numeros += ch
                row = int(numeros) if numeros else 1          # fila = dígitos
                col = self.letras_a_indice(letras)            # col = letras → índice
                if col > max_col_index:
                    max_col_index = col

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

                if row not in filas_dict:
                    filas_dict[row] = {}
                filas_dict[row][col] = val

            if not filas_dict:
                return []

            max_row = max(filas_dict.keys())
            matriz = []
            for r_i in range(1, max_row + 1):
                fila_vals = []
                row_dict = filas_dict.get(r_i, {})
                for c_j in range(1, max_col_index + 1):
                    fila_vals.append(row_dict.get(c_j, ""))
                matriz.append(fila_vals)

            return matriz

    def letras_a_indice(self, letters):
        letters = letters.upper()
        n = 0
        for ch in letters:
            n = n * 26 + (ord(ch) - ord('A') + 1)
        return n

# =================== Utilidades TEXTO (auto-separador) ===================

SEPARADORES_POSIBLES = [",", ";", "\t", "|", " "]

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
    print("\nSeparador sugerido:", repr("tabulador" if defecto == "\t" else ("espacio" if defecto == " " else defecto)))
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
    return [p.strip() for p in partes]

def lineas_a_tabla(lineas, sep):
    filas = []
    for l in lineas:
        if l.strip() == "":
            continue
        filas.append(partir_linea(l, sep))
    return filas

# =================== E/S segura y helpers numéricos ===================

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

def es_faltante(x):
    if x is None:
        return True
    t = to_lower(str(x).strip())
    return t in {"", "na", "nan", "null", "none"}

def intentar_float(x):
    try:
        return float(x)
    except Exception:
        return None

# =================== Lectura unificada ===================

def leer_matriz(ruta):
    # .xlsx o .csv directo vía lector
    if ruta.lower().endswith(".xlsx") or ruta.lower().endswith(".csv"):
        lector = XLSXtoCSV(ruta)
        return lector.procesar(), True  # True = ya es matriz (no pedir separador)
    # Texto general
    with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
        lineas = [ln.rstrip("\n\r") for ln in f if ln.strip() != ""]
    if not lineas:
        return [], True
    sep = pedir_separador(detectar_separador(lineas[:50]))
    return lineas_a_tabla(lineas, sep), True

def normalizar_ancho(filas):
    if not filas:
        return [], 0
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

# =================== Binarización robusta ===================

def a_binario(celda):
    if es_faltante(celda):
        return None
    s = to_lower(str(celda).strip())
    # Explícitos
    if s in {"1", "si", "sí", "true", "t", "y", "yes"}:
        return 1
    if s in {"0", "no", "false", "f", "n"}:
        return 0
    # Números
    f = intentar_float(s)
    if f is not None:
        return 1 if f != 0.0 else 0
    # No interpretable
    return None

# =================== Conteos A,B,C,D y coeficientes ===================

def contar_abcd(vec_i, vec_j):
    a=b=c=d=0
    for xi, xj in zip(vec_i, vec_j):
        if xi is None or xj is None:
            continue
        if xi==1 and xj==1:
            a += 1
        elif xi==1 and xj==0:
            b += 1
        elif xi==0 and xj==1:
            c += 1
        elif xi==0 and xj==0:
            d += 1
    return a,b,c,d

def matriz_abcd(matriz_bin):
    n = len(matriz_bin)
    A = [[0]*n for _ in range(n)]
    B = [[0]*n for _ in range(n)]
    C = [[0]*n for _ in range(n)]
    D = [[0]*n for _ in range(n)]
    for i in range(n):
        A[i][i]=B[i][i]=C[i][i]=0
        D[i][i]=sum(1 for v in matriz_bin[i] if v==0)  # opcional, diagonal informativa
        for j in range(i+1, n):
            a,b,c,d = contar_abcd(matriz_bin[i], matriz_bin[j])
            A[i][j]=A[j][i]=a
            B[i][j]=B[j][i]=b
            C[i][j]=C[j][i]=c
            D[i][j]=D[j][i]=d
    return A,B,C,D

def matriz_coeficientes(A,B,C,D):
    n = len(A)
    J = [[0.0]*n for _ in range(n)]   # Jaccard = a/(a+b+c)
    S = [[0.0]*n for _ in range(n)]   # Sokal–Michener = (a+d)/(a+b+c+d)
    for i in range(n):
        J[i][i]=1.0
        S[i][i]=1.0
        for j in range(i+1, n):
            a=A[i][j]; b=B[i][j]; c=C[i][j]; d=D[i][j]
            denom_J = a+b+c
            denom_S = a+b+c+d
            jacc = (float(a)/denom_J) if denom_J>0 else 0.0
            sok  = (float(a+d)/denom_S) if denom_S>0 else 0.0
            J[i][j]=J[j][i]=jacc
            S[i][j]=S[j][i]=sok
    return J,S

# =================== Impresión ===================

def imprimir_matriz(M, titulo, decimales=False):
    print("\n" + titulo)
    if not M:
        print("(vacío)")
        return
    n = len(M)
    ancho = 8
    print(" " * ancho, end="")
    for j in range(n):
        print(str(j+1).rjust(ancho), end="")
    print()
    for i in range(n):
        print(str(i+1).rjust(ancho), end="")
        for j in range(n):
            v = M[i][j]
            if decimales:
                try:
                    txt = "{:.4f}".format(float(v))
                except Exception:
                    txt = "0.0000"
            else:
                try:
                    txt = str(int(v))
                except Exception:
                    txt = "0"
            print(txt.rjust(ancho), end="")
        print()

# =================== Interfaz principal ===================

def main():
    print("=== Matrices A,B,C,D y coeficientes Jaccard / Sokal–Michener ===")
    ruta = seguro_input("Ruta del archivo (.xlsx o texto): ", default="")
    if not ruta:
        print("No se proporcionó ruta.")
        return

    try:
        filas = None
        if ruta.lower().endswith(".xlsx") or ruta.lower().endswith(".csv"):
            filas, _ = leer_matriz(ruta)
        else:
            filas, _ = leer_matriz(ruta)
    except Exception:
        print("No se pudo leer el archivo.")
        return

    if not filas:
        print("Sin datos.")
        return

    # ¿Encabezado?
    ans = seguro_input("¿La primera fila es encabezado? [s/n] (Enter = 's'): ", default="s")
    tiene_encabezado = (str(ans).strip().lower() != "n")

    filas_norm, ncols = normalizar_ancho(filas)
    if ncols == 0:
        print("Sin columnas.")
        return

    if tiene_encabezado:
        encabezado = filas_norm[0]
        datos = filas_norm[1:]
    else:
        encabezado = ["col_"+str(i+1) for i in range(ncols)]
        datos = filas_norm

    if not datos:
        print("Sin filas de datos.")
        return

    # Orientación
    print("\n¿Calcular por FILAS (individuos) o por COLUMNAS (variables)?")
    modo = None
    while modo not in {"f", "c"}:
        modo = seguro_input("Escribe 'f' para filas, 'c' para columnas: ", default="f")

    # Construir matriz binaria según orientación
    if modo == "f":
        # Cada fila es un vector binario sobre todas las columnas
        bin_rows = []
        for fila in datos:
            bin_rows.append([a_binario(x) for x in fila])
        # Filtramos columnas completamente no-binarias (todas None) para no sesgar
        # (no es obligatorio; solo evita columnas “basura”)
        col_ok = []
        for j in range(ncols):
            hay = False
            for i in range(len(bin_rows)):
                if bin_rows[i][j] is not None:
                    hay = True
                    break
            col_ok.append(hay)
        bin_rows = [ [v for v,ok in zip(row,col_ok) if ok] for row in bin_rows ]
        A,B,C,D = matriz_abcd(bin_rows)
        J,S = matriz_coeficientes(A,B,C,D)
        imprimir_matriz(A, "Matriz A (1-1)")
        imprimir_matriz(B, "Matriz B (1-0)")
        imprimir_matriz(C, "Matriz C (0-1)")
        imprimir_matriz(D, "Matriz D (0-0)")
        imprimir_matriz(J, "Matriz Jaccard", decimales=True)
        imprimir_matriz(S, "Matriz Sokal–Michener", decimales=True)
    else:
        # Por columnas: transponer primero (columna => vector)
        # Transposición segura
        nfil = len(datos)
        T = []
        for j in range(ncols):
            col = []
            for i in range(nfil):
                col.append(datos[i][j] if j < len(datos[i]) else "")
            T.append(col)
        bin_cols = []
        for col in T:
            bin_cols.append([a_binario(x) for x in col])
        # Filtrar filas totalmente None (opcional)
        row_ok = []
        for i in range(len(bin_cols[0]) if bin_cols else 0):
            hay = False
            for col in bin_cols:
                if col[i] is not None:
                    hay = True
                    break
            row_ok.append(hay)
        bin_cols = [ [v for v,ok in zip(col,row_ok) if ok] for col in bin_cols ]
        A,B,C,D = matriz_abcd(bin_cols)
        J,S = matriz_coeficientes(A,B,C,D)
        imprimir_matriz(A, "Matriz A (1-1) entre COLUMNAS")
        imprimir_matriz(B, "Matriz B (1-0) entre COLUMNAS")
        imprimir_matriz(C, "Matriz C (0-1) entre COLUMNAS")
        imprimir_matriz(D, "Matriz D (0-0) entre COLUMNAS")
        imprimir_matriz(J, "Matriz Jaccard entre COLUMNAS", decimales=True)
        imprimir_matriz(S, "Matriz Sokal–Michener entre COLUMNAS", decimales=True)

    print("\nListo.")

if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("Ocurrió un error no previsto. El programa finaliza de forma segura.")
