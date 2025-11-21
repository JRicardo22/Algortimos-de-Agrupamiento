try:
    # Cuando se ejecuta desde la carpeta raíz del proyecto
    from descompresor.Rms_lector import LectorXLSXCSV,LectorXLSXCSVError
except ImportError:
    # Posible ejecución como módulo dentro de un paquete
    from ..descompresor.Rms_lector import LectorXLSXCSVError

MAX_FILAS = 20000
MAX_COLUMNAS = 3000
MAX_PARES = 1_500_000

SEPARADORES_POSIBLES = [",", ";", "\t", "|", " "]

# -------- Utilidades generales --------

def seguro_input(mensaje, defecto=None):
    """
    Input "seguro": si el usuario solo presiona Enter y hay defecto, devuelve el defecto.
    Nunca lanza excepción hacia arriba.
    """
    while True:
        try:
            txt = input(mensaje)
        except Exception:
            # Si hay algún problema con stdin, usamos el defecto
            return defecto
        if txt == "" and defecto is not None:
            return defecto
        if txt is not None:
            return txt

def es_faltante(v):
    if v is None:
        return True
    t = str(v).strip()
    if t == "":
        return True
    t_low = t.lower()
    return t_low in ("na", "nan", "null", "none")

def to_lower(v):
    try:
        return str(v).strip().lower()
    except Exception:
        return ""

def intentar_float(v):
    try:
        if v is None:
            return None
        t = str(v).strip()
        if t == "":
            return None
        # Soportar coma decimal básica
        t = t.replace(",", ".")
        return float(t)
    except Exception:
        return None

def suma(valores):
    total = 0.0
    for v in valores:
        total += v
    return total

def minimo(valores):
    if not valores:
        return 0.0
    m = valores[0]
    for v in valores[1:]:
        if v < m:
            m = v
    return m

def maximo(valores):
    if not valores:
        return 0.0
    m = valores[0]
    for v in valores[1:]:
        if v > m:
            m = v
    return m

# -------- Lectura de archivos usando el descompresor --------

def leer_con_descompresor(ruta):
    """
    Usa XLSXtoCSV para leer .xlsx y .csv y devolver una lista de filas (lista de listas de str).
    """
    lector = leer_con_descompresor(ruta)
    filas = lector.procesar()
    # Aseguramos que todo sea str
    tabla = []
    for fila in filas:
        nueva = []
        for celda in fila:
            if celda is None:
                nueva.append("")
            else:
                nueva.append(str(celda))
        tabla.append(nueva)
    return tabla

# -------- Lectura de texto plano --------

def detectar_separador(lineas_muestra):
    """
    Intenta determinar el mejor separador entre varios candidatos.
    """
    mejor_sep = ","
    mejor_cols = 1
    mejor_consistencia = -1.0

    for sep in SEPARADORES_POSIBLES:
        conteos = []
        for ln in lineas_muestra:
            if sep == " ":
                partes = ln.strip().split()
            else:
                partes = ln.split(sep)
            if partes and any(p.strip() != "" for p in partes):
                conteos.append(len(partes))

        if not conteos:
            continue

        prom = suma(conteos) / float(len(conteos))
        var = 0.0
        for c in conteos:
            diff = c - prom
            var += diff * diff
        var = var / float(len(conteos))

        consistencia = 1.0 / (1.0 + var)
        cols_prom = prom

        if (consistencia > mejor_consistencia) or (
            abs(consistencia - mejor_consistencia) < 1e-9 and cols_prom > mejor_cols
        ):
            mejor_consistencia = consistencia
            mejor_cols = cols_prom
            mejor_sep = sep

    return mejor_sep

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

def leer_texto_a_tabla(ruta):
    with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
        lineas = [ln.rstrip("\n\r") for ln in f if ln.strip() != ""]
    if not lineas:
        raise Exception("El archivo de texto está vacío.")
    muestra = lineas[: min(30, len(lineas))]
    sep = detectar_separador(muestra)
    tabla = lineas_a_tabla(lineas, sep)
    return tabla

# -------- Lectura genérica de cualquier archivo tabular --------

def leer_tabla_desde_ruta():
    """
    Pide la ruta al usuario e intenta leer con el descompresor (.xlsx/.csv)
    o como texto plano, devolviendo (encabezado, datos).
    """
    while True:
        ruta = seguro_input("Ruta del archivo (.xlsx, .csv, .txt, etc.): ", defecto="")
        if not ruta:
            print("  No se proporcionó ruta.")
            continue

        ruta = ruta.strip()

        # Intento 1: descompresor
        tabla = None
        error_1 = None
        if ruta.lower().endswith(".xlsx") or ruta.lower().endswith(".csv"):
            try:
                tabla = leer_con_descompresor(ruta)
            except Exception as e:
                error_1 = e

        # Si no termina en .xlsx/.csv o falló el descompresor, probamos texto
        if tabla is None:
            try:
                tabla = leer_texto_a_tabla(ruta)
            except Exception as e:
                if error_1 is not None:
                    print("  Error usando el descompresor:", error_1)
                print("  No se pudo leer el archivo como texto:", e)
                print("  Verifica la ruta o el formato del archivo.")
                continue

        if not tabla or not tabla[0]:
            print("  El archivo no contiene datos válidos.")
            continue

        # Limitar tamaño
        nfilas = len(tabla)
        ncols = len(tabla[0])

        if nfilas > MAX_FILAS:
            print("  Aviso: se recortan las filas de {} a {} para evitar problemas.".format(nfilas, MAX_FILAS))
            tabla = tabla[:MAX_FILAS]
            nfilas = MAX_FILAS

        if ncols > MAX_COLUMNAS:
            print("  Aviso: se recortan las columnas de {} a {} para evitar problemas.".format(ncols, MAX_COLUMNAS))
            tabla = [fila[:MAX_COLUMNAS] for fila in tabla]
            ncols = MAX_COLUMNAS

        print("  Datos cargados: {} filas x {} columnas.".format(nfilas, ncols))

        usa_enc = seguro_input("¿La primera fila es encabezado? [S/n]: ", defecto="S")
        usa_enc = (usa_enc or "S").strip().lower()
        if usa_enc == "n":
            encabezado = ["Col{}".format(i + 1) for i in range(ncols)]
            datos = tabla
        else:
            encabezado = []
            for cel in tabla[0]:
                cel_txt = str(cel).strip()
                if cel_txt == "":
                    encabezado.append("Col{}".format(len(encabezado) + 1))
                else:
                    encabezado.append(cel_txt)
            datos = tabla[1:]
            if not datos:
                print("  Solo había una fila (el encabezado). No hay datos.")
                continue

        return encabezado, datos

# -------- Tipificación de columnas y rangos numéricos --------

def tipificar_columnas(datos):
    """
    Clasifica cada columna como:
        - 'numerico'
        - 'binario_numerico'
        - 'binario_categorico'
        - 'categorico'
    """
    if not datos:
        return []

    ncols = len(datos[0])
    tipos = []

    for c in range(ncols):
        valores_crudos = [fila[c] for fila in datos]
        valores_validos = [v for v in valores_crudos if not es_faltante(v)]

        if not valores_validos:
            # Sin información: lo tratamos como categórico
            tipos.append("categorico")
            continue

        # Intento numérico
        nums = []
        num_valid = 0
        for v in valores_validos:
            fv = intentar_float(v)
            if fv is not None:
                nums.append(fv)
                num_valid += 1

        total = len(valores_validos)
        es_num = (total > 0 and num_valid >= max(1, int(0.8 * total)))

        if es_num:
            # ¿Es binario numérico (por ejemplo 0/1, 1/2, etc.)?
            unicos = []
            for x in nums:
                if x not in unicos:
                    unicos.append(x)
            if len(unicos) == 2:
                tipos.append("binario_numerico")
            else:
                tipos.append("numerico")
        else:
            # Lo tratamos como categórico y vemos si sólo hay 2 categorías
            unicos = []
            for v in valores_validos:
                lv = to_lower(v)
                if lv not in unicos:
                    unicos.append(lv)
            if len(unicos) == 2:
                tipos.append("binario_categorico")
            else:
                tipos.append("categorico")

    return tipos

def rangos_numericos(datos, tipos):
    """
    Calcula (min, max) para las columnas numéricas; para las no numéricas pone (0, 0).
    """
    ncols = len(tipos)
    rangos = []
    for c in range(ncols):
        if tipos[c].startswith("numerico") or tipos[c] == "binario_numerico":
            vals = []
            for fila in datos:
                if es_faltante(fila[c]):
                    continue
                fv = intentar_float(fila[c])
                if fv is not None:
                    vals.append(fv)
            if not vals:
                rangos.append((0.0, 0.0))
            else:
                rangos.append((minimo(vals), maximo(vals)))
        else:
            rangos.append((0.0, 0.0))
    return rangos

# -------- Núcleo de Gower --------

def similitud_gower_registro(a, b, tipos, rangos):
    """
    Calcula similitud de Gower entre dos registros (filas).
    Devuelve (s, d, k) donde:
        s = similitud
        d = distancia = 1 - s
        k = número de columnas efectivamente comparadas
    """
    n = min(len(a), len(b), len(tipos))
    num = 0.0
    den = 0.0
    k_activos = 0

    for c in range(n):
        tipo = tipos[c]
        va = a[c]
        vb = b[c]

        if es_faltante(va) or es_faltante(vb):
            continue

        if tipo.startswith("numerico"):
            xa = intentar_float(va)
            xb = intentar_float(vb)
            if xa is None or xb is None:
                continue
            mn, mx = rangos[c]
            if mx > mn:
                s_ijk = 1.0 - abs(xa - xb) / float(mx - mn)
            else:
                # Sin rango (todos iguales): máxima similitud
                s_ijk = 1.0
        else:
            # Para categórico y binario: coincidencia 1 / 0
            s_ijk = 1.0 if to_lower(va) == to_lower(vb) else 0.0

        num += s_ijk
        den += 1.0
        k_activos += 1

    if den == 0.0:
        # No se pudo comparar ninguna columna
        return 0.0, 1.0, 0

    s = num / den
    d = 1.0 - s
    return s, d, k_activos

def matriz_completa(datos, tipos, rangos):
    """
    Construye la matriz de similitud y distancia entre TODAS las filas.
    """
    n = len(datos)
    if n <= 1:
        return [], []

    pares = n * (n - 1) // 2
    if pares > MAX_PARES:
        print("  Aviso: hay demasiados pares ({}). Se calcularán solo los primeros {}.".format(pares, MAX_PARES))

    S = [[1.0 if i == j else 0.0 for j in range(n)] for i in range(n)]
    D = [[0.0 if i == j else 1.0 for j in range(n)] for i in range(n)]

    cuenta = 0
    for i in range(n):
        for j in range(i + 1, n):
            if cuenta >= MAX_PARES:
                break
            s, d, _ = similitud_gower_registro(datos[i], datos[j], tipos, rangos)
            S[i][j] = S[j][i] = s
            D[i][j] = D[j][i] = d
            cuenta += 1
        if cuenta >= MAX_PARES:
            break

    return S, D

# -------- Utilidades de impresión --------

def formato_float(x, ancho=7, decimales=4):
    try:
        txt = ("{0:." + str(decimales) + "f}").format(float(x))
    except Exception:
        txt = "0.0"
    if len(txt) < ancho:
        txt = " " * (ancho - len(txt)) + txt
    return txt

def imprimir_par_s_d(s, d, etiqueta_i, etiqueta_j, k):
    print("")
    print("Par de registros:")
    print("  A =", etiqueta_i)
    print("  B =", etiqueta_j)
    print("  Columnas comparadas (válidas):", k)
    print("  Similitud s =", formato_float(s))
    print("  Distancia d = 1 - s =", formato_float(d))

def imprimir_matriz(M, titulo):
    if not M:
        print("  (matriz vacía)")
        return
    n = len(M)
    print("")
    print(titulo)
    for i in range(n):
        fila = []
        for j in range(len(M[i])):
            fila.append(formato_float(M[i][j]))
        print(" ".join(fila))

def pedir_indice(n, nombre="fila"):
    while True:
        txt = seguro_input("Índice de {} (1..{}): ".format(nombre, n), defecto="")
        if not txt:
            print("  Entrada vacía.")
            continue
        try:
            idx = int(txt)
            if 1 <= idx <= n:
                return idx - 1
        except Exception:
            pass
        print("  Índice inválido, intenta de nuevo.")

def pedir_par(n, nombre="fila"):
    while True:
        txt = seguro_input("Par de {} (i,j) con 1..{} separados por coma: ".format(nombre, n), defecto="")
        if not txt:
            print("  Entrada vacía.")
            continue
        trozos = txt.replace("(", "").replace(")", "").split(",")
        if len(trozos) != 2:
            print("  Debes escribir dos índices separados por coma, por ejemplo 1,3.")
            continue
        try:
            a = int(trozos[0].strip())
            b = int(trozos[1].strip())
            if 1 <= a <= n and 1 <= b <= n and a != b:
                return a - 1, b - 1
        except Exception:
            pass
        print("  Índices inválidos o iguales.")

# -------- Programa principal --------

def main():
    print("=== PROGRAMA 4: Gower robusto con descompresor ===")

    # 1) Leer datos
    encabezado, datos = leer_tabla_desde_ruta()

    # 2) Tipificar columnas y calcular rangos
    tipos = tipificar_columnas(datos)
    rangos = rangos_numericos(datos, tipos)

    print("")
    print("Tipos de columnas detectados:")
    for i, (nom, t) in enumerate(zip(encabezado, tipos), start=1):
        print("  Col {:3d}: {:20s} -> {}".format(i, str(nom)[:20], t))

    # 3) Elegir operación
    print("")
    print("¿Qué deseas calcular?")
    print("  1) Similitud/distancia entre DOS FILAS")
    print("  2) Matriz completa entre TODAS las FILAS")
    op = seguro_input("Elige opción [1/2]: ", defecto="1")
    if op not in ("1", "2"):
        op = "1"

    nfilas = len(datos)

    if op == "1":
        print("")
        print("Se calculará Gower para un par de filas (registros).")
        i, j = pedir_par(nfilas, "fila")
        s, d, k = similitud_gower_registro(datos[i], datos[j], tipos, rangos)
        et_i = "Fila {}".format(i + 1)
        et_j = "Fila {}".format(j + 1)
        imprimir_par_s_d(s, d, et_i, et_j, k)
    else:
        print("")
        print("Se calculará la matriz completa entre todas las filas.")
        S, D = matriz_completa(datos, tipos, rangos)
        imprimir_matriz(S, "Matriz de SIMILITUD entre FILAS (s)")
        imprimir_matriz(D, "Matriz de DISTANCIA entre FILAS (d = 1 - s)")

    print("")
    print(" Listo, el programa terminó sin errores graves.")

if __name__ == "__main__":
    try:
        main()
    except Exception:
        # Nunca dejamos escapar un trace completo para que no se rompa
        print("  Ocurrió un error no previsto. El programa finaliza de forma segura.")
