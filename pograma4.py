# -*- coding: utf-8 -*-
"""
Gower PURO con soporte XLSX (sin librerías externas).
- Solo usa biblioteca estándar: zipfile, xml.etree.ElementTree (sin os, sin sys).
- Lee .xlsx (Excel): toma la HOJA 1 por defecto (o de 1..N a elección).
- Lee texto: .txt, .csv, .tsv, .pipe, etc. (auto-detección de separador).
- Calcula similitud (s) y distancia (d = 1 - s) de Gower:
    * Por FILAS (individuos) o por COLUMNAS (variables).
    * Para un PAR de índices o MATRIZ COMPLETA (con límites).
- Manejo de faltantes: "", NA, NaN, NULL, None (may/minus).

Límites (para no colgar la compu):
- Máx. filas: 20000
- Máx. columnas: 3000
- Máx. pares matriz: 1_500_000
"""

from zipfile import ZipFile
from xml.etree import ElementTree

# ----------------- Parámetros de seguridad -----------------
MAX_FILAS = 20000
MAX_COLUMNAS = 3000
MAX_PARES_MATRIZ = 1_500_000

FALTANTES = {"", "na", "nan", "null", "none"}  # se usa .strip().lower()
SEPARADORES_POSIBLES = [",", ";", "\t", "|", " "]

# ----------------- Utilidades de entrada -------------------

def seguro_input(prompt, default=None, validar=None):
    while True:
        try:
            txt = input(prompt)
        except Exception:
            return default
        if validar is None or validar(txt):
            return txt
        print("  Entrada inválida, intenta de nuevo.")

def leer_ruta_y_tipo():
    """
    Pide la ruta y detecta si es XLSX (zip con [Content_Types].xml) o texto plano.
    Sin os/sys: validamos abriendo.
    """
    while True:
        ruta = seguro_input("Ruta del archivo (.xlsx o texto): ", default="")
        if not ruta:
            print("  No se proporcionó ruta.")
            continue
        # ¿Es XLSX?
        try:
            with ZipFile(ruta, "r") as z:
                # Validación mínima de XLSX
                if "[Content_Types].xml" in z.namelist():
                    return ruta, "xlsx"
        except Exception:
            pass
        # Si no es XLSX, intentamos abrir como texto
        try:
            with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
                lineas = [ln.rstrip("\n\r") for ln in f if ln.strip() != ""]
            if not lineas:
                print(" El archivo está vacío o no es válido como texto.")
                continue
            return ruta, "texto"
        except Exception:
            print(" No se pudo abrir/leer el archivo como XLSX ni como texto. Verifica la ruta.")
            continue

# --------------- Lectura TEXTO (CSV/TSV/etc.) ---------------

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
    print("  Entrada no válida; usando el sugerido.")
    return defecto

def partir_linea(linea, sep):
    if sep == " ":
        partes = linea.strip().split()
    else:
        partes = linea.split(sep)
    limpiado = []
    for p in partes:
        limpiado.append(p.strip())
    return limpiado

def lineas_a_tabla(lineas, sep):
    filas = []
    for l in lineas:
        if l.strip() == "":
            continue
        fila = partir_linea(l, sep)
        filas.append(fila)
    return filas

# ----------- Lectura XLSX (sin dependencias externas) -----------

def leer_xlsx_a_matriz(ruta, sheet_index=None):
    """
    Convierte un .xlsx a una matriz de strings (lista de filas).
    - sheet_index: índice 1..N (por defecto toma 1).
    - Convierte strings compartidas (sharedStrings). Otras celdas como texto del <v>.
    - Fechas/estilos: se dejan en crudo (número de Excel) para mantener pureza sin mapear estilos.
    """
    with ZipFile(ruta, "r") as z:
        names = z.namelist()

        # Shared Strings (opcional)
        shared = []
        if "xl/sharedStrings.xml" in names:
            try:
                sst = ElementTree.fromstring(z.read("xl/sharedStrings.xml"))
                # namespace no siempre declarado; usamos búsqueda genérica
                for si in sst.iter():
                    if si.tag.endswith("si"):
                        # concatenamos todos los textos dentro de si
                        txt = ""
                        for tnode in si.iter():
                            if tnode.tag.endswith("t") and tnode.text is not None:
                                txt += tnode.text
                        shared.append(txt)
            except Exception:
                shared = []

        # Workbook para conocer hojas
        sheets = []
        try:
            wb = ElementTree.fromstring(z.read("xl/workbook.xml"))
            for node in wb.iter():
                if node.tag.endswith("sheet"):
                    # nombre y r:id (no lo usamos aquí)
                    nm = node.attrib.get("name", "")
                    sheets.append(nm)
        except Exception:
            pass

        # Determinar hoja
        total = len([n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")])
        if total == 0:
            raise ValueError("XLSX no contiene hojas legibles.")
        if sheet_index is None:
            # preguntamos si hay más de una
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
            # fallback por si el orden físico difiere
            # intentamos el primero disponible
            cand = [n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
            cand.sort()
            sheet_path = cand[0]

        # Parsear celdas
        xml = ElementTree.fromstring(z.read(sheet_path))
        # Construiremos un dict por fila -> dict col_index -> valor
        filas_dict = {}
        max_col_index = 0

        for c in xml.iter():
            if not c.tag.endswith("c"):
                continue
            r = c.attrib.get("r", "")  # ej. "C5"
            t = c.attrib.get("t", "")  # tipo (s = shared string, b = bool, etc.)
            v_node = None
            for sub in c.iter():
                if sub.tag.endswith("v"):
                    v_node = sub
                    break
            if not r:
                continue
            fila_idx, col_idx = ref_a_indices(r)  # 1-based -> (row, col)
            if col_idx > max_col_index:
                max_col_index = col_idx
            val = ""
            if v_node is not None and v_node.text is not None:
                raw = v_node.text
                if t == "s":
                    # shared string
                    try:
                        si = int(raw)
                        if 0 <= si < len(shared):
                            val = shared[si]
                        else:
                            val = raw
                    except Exception:
                        val = raw
                elif t == "b":
                    # boolean
                    val = "1" if raw.strip() == "1" else "0"
                else:
                    # numérico o texto directo
                    val = raw
            else:
                # celdas sin <v> se pueden considerar vacías
                val = ""

            if fila_idx not in filas_dict:
                filas_dict[fila_idx] = {}
            filas_dict[fila_idx][col_idx] = val

        # Convertir dict disperso a matriz compacta (sin filas vacías al final)
        if not filas_dict:
            return []

        max_row = max(filas_dict.keys())
        matriz = []
        for r_i in range(1, max_row + 1):
            fila_vals = []
            row_dict = filas_dict.get(r_i, {})
            # si la fila está totalmente vacía (sin claves), podemos colocar []
            if not row_dict:
                # Aun así mantenemos la estructura: fila vacía
                # Si no deseas filas vacías, podrías "continue"
                fila_vals = ["" for _ in range(max_col_index)]
            else:
                for c_j in range(1, max_col_index + 1):
                    fila_vals.append(row_dict.get(c_j, ""))
            matriz.append(fila_vals)

        return matriz

# ----------- Helpers XLSX: referencias y letras de columna -----------

def ref_a_indices(ref):
    """
    Convierte una referencia tipo "C5" a (row=5, col=3), ambos 1-based.
    """
    letras = ""
    numeros = ""
    for ch in ref:
        if "A" <= ch <= "Z" or "a" <= ch <= "z":
            letras += ch
        elif "0" <= ch <= "9":
            numeros += ch
    col = letras_a_indice(letras)
    row = int(numeros) if numeros else 1
    return row, col

def letras_a_indice(letters):
    """
    'A'->1, 'B'->2, ..., 'Z'->26, 'AA'->27, etc.
    """
    letters = letters.upper()
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

# ----------------- Utilidades numéricas puras -----------------

def suma(valores):
    total = 0.0
    for v in valores:
        total += v
    return total

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

def es_faltante(celda):
    if celda is None:
        return True
    t = to_lower(celda.strip())
    return t in FALTANTES

def to_lower(s):
    try:
        return s.lower()
    except Exception:
        return s

def intentar_float(txt):
    try:
        return float(txt)
    except Exception:
        return None

def normalizar_ancho(filas):
    """
    Ajusta todas las filas al ancho modal (la cantidad de columnas más repetida).
    """
    conteo_por_long = {}
    for fila in filas:
        n = len(fila)
        conteo_por_long[n] = conteo_por_long.get(n, 0) + 1
    ncols_obj = None
    mejor_freq = -1
    for k, freq in conteo_por_long.items():
        if freq > mejor_freq:
            mejor_freq = freq
            ncols_obj = k
    filas_norm = []
    for fila in filas:
        if len(fila) < ncols_obj:
            fila = fila + [""] * (ncols_obj - len(fila))
        elif len(fila) > ncols_obj:
            fila = fila[:ncols_obj]
        filas_norm.append(fila)
    return filas_norm, ncols_obj

# ----------------- Tipificación y rangos -----------------

def tipificar_columnas(datos):
    if not datos:
        return []
    ncols = len(datos[0])
    tipos = []
    for c in range(ncols):
        es_num = True
        for fila in datos:
            v = fila[c]
            if es_faltante(v):
                continue
            if intentar_float(v) is None:
                es_num = False
                break
        tipos.append("numerico" if es_num else "categorico")
    return tipos

def rangos_numericos(datos, tipos):
    if not datos:
        return []
    ncols = len(datos[0])
    r = []
    for c in range(ncols):
        if tipos[c] != "numerico":
            r.append(None)
            continue
        vals = []
        for fila in datos:
            v = fila[c]
            if es_faltante(v):
                continue
            f = intentar_float(v)
            if f is not None:
                vals.append(f)
        if not vals:
            r.append((0.0, 0.0))
        else:
            r.append((minimo(vals), maximo(vals)))
    return r

# ----------------- Núcleo de Gower -----------------

def similitud_gower_registro(a, b, tipos, rangos):
    try:
        num = 0.0
        den = 0.0
        ncols = len(tipos)
        k = 0
        while k < ncols:
            va = a[k]; vb = b[k]
            if (not es_faltante(va)) and (not es_faltante(vb)):
                if tipos[k] == "numerico":
                    fa = intentar_float(va); fb = intentar_float(vb)
                    if (fa is not None) and (fb is not None):
                        r = rangos[k]
                        if r is None:
                            mn = 0.0; mx = 0.0
                        else:
                            mn, mx = r[0], r[1]
                        R = mx - mn
                        if R == 0.0:
                            s_k = 1.0
                        else:
                            dif = fa - fb
                            if dif < 0: dif = -dif
                            s_k = 1.0 - (dif / R)
                            if s_k < 0.0: s_k = 0.0
                            if s_k > 1.0: s_k = 1.0
                        num += s_k
                        den += 1.0
                else:
                    s_k = 1.0 if va == vb else 0.0
                    num += s_k
                    den += 1.0
            k += 1
        if den == 0.0:
            return 0.0, 1.0, 0
        s = num / den
        d = 1.0 - s
        if s < 0.0: s = 0.0
        if s > 1.0: s = 1.0
        if d < 0.0: d = 0.0
        if d > 1.0: d = 1.0
        return s, d, int(den)
    except Exception:
        return 0.0, 1.0, 0

def matriz_completa(datos, tipos, rangos):
    n = len(datos)
    pares = (n * (n - 1)) // 2
    if pares > MAX_PARES_MATRIZ:
        print(" Matriz demasiado grande:", n, "→", pares, "(límite:", MAX_PARES_MATRIZ, ").")
        print("   Sugerencia: calcule por PAR o reduzca el tamaño.")
        return None, None
    S = []
    D = []
    i = 0
    while i < n:
        fila_s = []
        fila_d = []
        j = 0
        while j < n:
            if i == j:
                fila_s.append(1.0)
                fila_d.append(0.0)
            elif j < i:
                fila_s.append(S[j][i])
                fila_d.append(D[j][i])
            else:
                s, d, _ = similitud_gower_registro(datos[i], datos[j], tipos, rangos)
                fila_s.append(s)
                fila_d.append(d)
            j += 1
        S.append(fila_s)
        D.append(fila_d)
        i += 1
    return S, D

def transponer(matriz):
    if not matriz:
        return []
    nfil = len(matriz)
    ncol = len(matriz[0])
    T = []
    j = 0
    while j < ncol:
        fila = []
        i = 0
        while i < nfil:
            fila.append(matriz[i][j])
            i += 1
        T.append(fila)
        j += 1
    return T

# ----------------- Impresión -----------------

def imprimir_par_s_d(s, d, etiqueta_a, etiqueta_b, k_usables):
    print("\n================= RESULTADO =================")
    print("Comparación:", etiqueta_a, " vs ", etiqueta_b)
    print("Atributos usados (no vacíos en ambos):", k_usables)
    print("Similitud de Gower (s):", formato_float(s))
    print("Distancia de Gower (d = 1 - s):", formato_float(d))
    print("============================================\n")

def imprimir_matriz(M, titulo):
    if M is None:
        return
    n = len(M)
    ancho = 10
    print("\n" + titulo + " (primeros índices como guía):")
    print(" " * ancho, end="")
    j = 0
    while j < n:
        txt = str(j + 1)
        print(txt.rjust(ancho), end="")
        j += 1
    print()
    i = 0
    while i < n:
        txti = str(i + 1)
        print(txti.rjust(ancho), end="")
        fila = M[i]
        j = 0
        while j < n:
            val = fila[j]
            try:
                val = float(val)
            except Exception:
                val = 0.0
            s = formatea_ancho(val, ancho)
            print(s, end="")
            j += 1
        print()
        i += 1
    print()

def formato_float(x):
    try:
        return "{:.6f}".format(float(x))
    except Exception:
        return "0.000000"

def formatea_ancho(x, ancho):
    try:
        return ("{:" + str(ancho) + ".6f}").format(float(x))
    except Exception:
        return ("{:" + str(ancho) + "s}").format("0.000000")

# ----------------- Flujo principal -----------------

def main():
    print("=== Similitud/Distancia de Gower==")
    ruta, tipo = leer_ruta_y_tipo()

    if tipo == "xlsx":
        try:
            matriz = leer_xlsx_a_matriz(ruta, sheet_index=None)
        except Exception:
            print(" No se pudo leer el XLSX. Verifica que no esté corrupto.")
            return
        if not matriz:
            print(" El XLSX no aportó datos.")
            return

        # ¿Encabezado?
        ans = seguro_input("¿La primera fila es encabezado? [s/n] (Enter = 's'): ", default="s")
        tiene_encabezado = (str(ans).strip().lower() != "n")

        filas_norm, ncols_obj = normalizar_ancho(matriz)
        if tiene_encabezado:
            encabezado = filas_norm[0]
            datos = filas_norm[1:]
        else:
            encabezado = []
            i = 0
            while i < ncols_obj:
                encabezado.append("col_" + str(i + 1))
                i += 1
            datos = filas_norm

    else:
        # TEXTO
        try:
            with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
                lineas = [ln.rstrip("\n\r") for ln in f if ln.strip() != ""]
        except Exception:
            print(" No se pudo abrir/leer el archivo de texto.")
            return
        muestra = lineas[:50]
        sep_detectado = detectar_separador(muestra)
        sep = pedir_separador(sep_detectado)
        ans = seguro_input("¿La primera fila es encabezado? [s/n] (Enter = 's'): ", default="s")
        tiene_encabezado = (str(ans).strip().lower() != "n")
        filas = lineas_a_tabla(lineas, sep)
        filas_norm, ncols_obj = normalizar_ancho(filas)
        if tiene_encabezado:
            encabezado = filas_norm[0]
            datos = filas_norm[1:]
        else:
            encabezado = []
            i = 0
            while i < ncols_obj:
                encabezado.append("col_" + str(i + 1))
                i += 1
            datos = filas_norm

    # Límites de seguridad
    nfil = len(datos); ncol = len(encabezado)
    if nfil > MAX_FILAS:
        print(" Demasiadas filas:", nfil, "(límite:", MAX_FILAS, ").")
        print("   Reduce el archivo.")
        return
    if ncol > MAX_COLUMNAS:
        print(" Demasiadas columnas:", ncol, "(límite:", MAX_COLUMNAS, ").")
        print("   Reduce el archivo.")
        return

    # Elección de orientación y operación
    print("\n¿Deseas calcular por FILAS (individuos) o por COLUMNAS (variables)?")
    modo = None
    while modo not in {"f", "c"}:
        modo = seguro_input("Escribe 'f' para filas, 'c' para columnas: ", default="")
        if modo not in {"f", "c"}:
            print("  Opción inválida.")

    print("\n¿Qué deseas calcular?")
    print("  1) s y d para un PAR específico")
    print("  2) MATRIZ COMPLETA s y d")
    op = None
    while op not in {"1", "2"}:
        op = seguro_input("Elige 1 o 2: ", default="")
        if op not in {"1", "2"}:
            print("  Opción inválida.")

    if modo == "f":
        tipos = tipificar_columnas(datos)
        rangos = rangos_numericos(datos, tipos)
        if op == "1":
            i, j = pedir_par(len(datos), "fila")
            s, d, k = similitud_gower_registro(datos[i], datos[j], tipos, rangos)
            imprimir_par_s_d(s, d, "Fila " + str(i + 1), "Fila " + str(j + 1), k)
        else:
            S, D = matriz_completa(datos, tipos, rangos)
            imprimir_matriz(S, "Matriz de SIMILITUD (s)")
            imprimir_matriz(D, "Matriz de DISTANCIA (d = 1 - s)")
    else:
        T = transponer(datos)
        tipos_col = tipificar_columnas(T)
        rangos_col = rangos_numericos(T, tipos_col)
        if op == "1":
            i, j = pedir_par(len(T), "columna")
            s, d, k = similitud_gower_registro(T[i], T[j], tipos_col, rangos_col)
            et_i = encabezado[i] if i < len(encabezado) else "Col " + str(i + 1)
            et_j = encabezado[j] if j < len(encabezado) else "Col " + str(j + 1)
            imprimir_par_s_d(s, d, et_i, et_j, k)
        else:
            S, D = matriz_completa(T, tipos_col, rangos_col)
            imprimir_matriz(S, "Matriz de SIMILITUD entre COLUMNAS (s)")
            imprimir_matriz(D, "Matriz de DISTANCIA entre COLUMNAS (d = 1 - s)")

    print("✓ Listo.")

# -------- Pequeñas utilidades finales --------

def pedir_par(n, etiqueta):
    while True:
        a_txt = seguro_input("Elige " + etiqueta + " A (1.." + str(n) + "): ", default="")
        b_txt = seguro_input("Elige " + etiqueta + " B (1.." + str(n) + "): ", default="")
        try:
            a = int(a_txt)
            b = int(b_txt)
            if 1 <= a <= n and 1 <= b <= n and a != b:
                return a - 1, b - 1
        except Exception:
            pass
        print("  Índices inválidos o iguales.")

# ----------------- Lanzador -----------------

if __name__ == "__main__":
    try:
        main()
    except Exception:
        # Salida limpia y sin stacktrace
        print("  Ocurrió un error no previsto. El programa finaliza de forma segura.")
