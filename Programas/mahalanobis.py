from descompresor.lee import XLSXtoCSV   # <-- AJUSTA ESTE NOMBRE

# FUNCIONES DE ÁLGEBRA
def a_float_seguro(cadena):
    """
    Intenta convertir una cadena a float.
    Soporta:
      - cadenas vacías -> error
      - números con punto
      - números con coma como separador decimal
    """
    if cadena is None:
        raise ValueError("Valor None")
    s = str(cadena).strip()
    if s == "":
        raise ValueError("Cadena vacía")
    # Intento directo
    try:
        return float(s)
    except ValueError:
        # Intento reemplazando coma por punto
        if "," in s:
            try:
                return float(s.replace(",", "."))
            except ValueError:
                pass
        raise ValueError(f"No se puede convertir a número: {repr(cadena)}")


def calcular_media(datos):
    if not datos:
        raise ValueError("No hay datos para calcular la media.")
    n_muestras = len(datos)
    n_dim = len(datos[0])
    media = [0.0] * n_dim

    for j in range(n_dim):
        s = 0.0
        for i in range(n_muestras):
            s += datos[i][j]
        media[j] = s / n_muestras
    return media


def calcular_matriz_covarianza(datos, media):
    n_muestras = len(datos)
    if n_muestras < 2:
        raise ValueError("Se requieren al menos 2 filas válidas para calcular la covarianza.")

    n_dim = len(datos[0])
    cov = [[0.0 for _ in range(n_dim)] for _ in range(n_dim)]

    for j in range(n_dim):
        for k in range(j, n_dim):
            suma = 0.0
            for i in range(n_muestras):
                suma += (datos[i][j] - media[j]) * (datos[i][k] - media[k])
            valor = suma / (n_muestras - 1)
            cov[j][k] = valor
            cov[k][j] = valor
    return cov


def matriz_identidad(n):
    I = [[0.0 for _ in range(n)] for _ in range(n)]
    for i in range(n):
        I[i][i] = 1.0
    return I


def inversa_matriz(A):
    """
    Inversa por Gauss-Jordan.
    Lanza ValueError si la matriz es singular o casi singular.
    """
    n = len(A)
    if n == 0 or any(len(f) != n for f in A):
        raise ValueError("La matriz debe ser cuadrada y no vacía.")

    M = [[float(A[i][j]) for j in range(n)] for i in range(n)]
    Inv = matriz_identidad(n)

    for col in range(n):
        # Buscar pivote
        pivote = col
        for fila in range(col + 1, n):
            if abs(M[fila][col]) > abs(M[pivote][col]):
                pivote = fila

        if abs(M[pivote][col]) < 1e-12:
            raise ValueError("La matriz es singular o casi singular (no invertible).")

        # Intercambiar filas si hace falta
        if pivote != col:
            M[col], M[pivote] = M[pivote], M[col]
            Inv[col], Inv[pivote] = Inv[pivote], Inv[col]

        # Normalizar fila pivote
        diag = M[col][col]
        for j in range(n):
            M[col][j] /= diag
            Inv[col][j] /= diag

        # Eliminar en otras filas
        for fila in range(n):
            if fila != col:
                factor = M[fila][col]
                for j in range(n):
                    M[fila][j] -= factor * M[col][j]
                    Inv[fila][j] -= factor * Inv[col][j]

    return Inv


def multiplicar_matriz_vector(A, v):
    n_filas = len(A)
    if n_filas == 0:
        raise ValueError("Matriz vacía en multiplicar_matriz_vector.")
    n_cols = len(A[0])
    if len(v) != n_cols:
        raise ValueError("Dimensiones incompatibles en multiplicar_matriz_vector.")
    resultado = [0.0] * n_filas
    for i in range(n_filas):
        s = 0.0
        for j in range(n_cols):
            s += A[i][j] * v[j]
        resultado[i] = s
    return resultado


def producto_punto(u, v):
    if len(u) != len(v):
        raise ValueError("Dimensiones incompatibles en producto_punto.")
    s = 0.0
    for i in range(len(u)):
        s += u[i] * v[i]
    return s


def distancia_mahalanobis(punto, media, cov_inv):
    if len(punto) != len(media):
        raise ValueError("Dimensiones incompatibles entre punto y media.")
    d = [punto[i] - media[i] for i in range(len(punto))]
    w = multiplicar_matriz_vector(cov_inv, d)
    valor = producto_punto(d, w)
    if valor < 0 and abs(valor) < 1e-12:
        valor = 0.0
    if valor < 0:
        raise ValueError("Valor interno negativo en Mahalanobis (revisa datos).")
    return valor ** 0.5

def cargar_filas_desde_archivo(ruta):
    """
     XLSXtoCSV para leer .csv, .txt, .xlsx, etc.
    Luego TRASPONE lo que devuelve para que queden filas normales:
      - cada fila = una observación
      - cada columna = una variable
    """
    if not isinstance(ruta, str) or ruta.strip() == "":
        raise ValueError("Ruta de archivo vacía o inválida.")

    try:
        conv = XLSXtoCSV(ruta)
        columnas = conv.procesar()  # OJO: esto vienen como columnas
    except FileNotFoundError:
        raise FileNotFoundError(f"No se encontró el archivo: {ruta}")
    except PermissionError:
        raise PermissionError(f"Permiso denegado al intentar leer: {ruta}")
    except Exception as e:
        raise Exception(f"Error al procesar el archivo con XLSXtoCSV: {e}")

    if not columnas or not isinstance(columnas, list):
        raise ValueError("El archivo no contiene datos válidos.")

    # Asegurar que cada "columna" sea lista
    cols_limpias = []
    for col in columnas:
        if col is None:
            continue
        if isinstance(col, list):
            cols_limpias.append([("" if c is None else str(c)) for c in col])
        else:
            cols_limpias.append([str(col)])

    if not cols_limpias:
        raise ValueError("No se obtuvieron columnas válidas del archivo.")

    # Trasponer: columnas -> filas
    n_cols = len(cols_limpias)
    n_rows = len(cols_limpias[0])
    for c in cols_limpias:
        if len(c) != n_rows:
            raise ValueError("Las columnas tienen longitudes distintas; archivo inconsistente.")

    filas = []
    for i in range(n_rows):
        fila = []
        for j in range(n_cols):
            fila.append(cols_limpias[j][i])
        filas.append(fila)

    return filas


def seleccionar_columnas(filas):
    if not filas:
        raise ValueError("No hay filas para seleccionar columnas.")

    max_cols = max(len(f) for f in filas)
    if max_cols == 0:
        raise ValueError("Las filas no contienen columnas.")

    print(f"\nEl archivo tiene {max_cols} columnas (numeradas desde 1).")
    print("Ejemplo de la primera fila:")
    print(filas[0])

    while True:
        entrada = input("Ingresa los números de las columnas numéricas, separados por comas (ej. 1,3,4): ").strip()
        if not entrada:
            print("Debes ingresar al menos una columna.")
            continue
        partes = entrada.split(",")
        indices = []
        valido = True
        for p in partes:
            p = p.strip()
            if not p.isdigit():
                print(f"Valor no numérico en índices de columnas: {repr(p)}")
                valido = False
                break
            col_num = int(p)
            if col_num < 1 or col_num > max_cols:
                print(f"Índice de columna fuera de rango: {col_num}")
                valido = False
                break
            indices.append(col_num - 1)
        if not valido:
            continue
        if not indices:
            print("No se seleccionó ninguna columna.")
            continue
        return indices


def construir_matriz_numerica(filas, usar_encabezado, indices_columnas):
    if not filas:
        raise ValueError("No hay filas para construir la matriz numérica.")

    inicio = 1 if usar_encabezado and len(filas) > 1 else 0
    datos = []
    filas_invalidas = 0

    for idx_fila in range(inicio, len(filas)):
        fila = filas[idx_fila]
        if fila is None:
            filas_invalidas += 1
            continue

        fila_numerica = []
        error_en_fila = False
        for col_idx in indices_columnas:
            if col_idx >= len(fila):
                error_en_fila = True
                break
            valor_str = fila[col_idx]
            try:
                valor_float = a_float_seguro(valor_str)
                fila_numerica.append(valor_float)
            except ValueError:
                error_en_fila = True
                break

        if error_en_fila:
            filas_invalidas += 1
            continue

        datos.append(fila_numerica)

    total_entrada = len(filas) - inicio
    if not datos:
        raise ValueError("Todas las filas fueron inválidas; no hay datos numéricos suficientes.")
    if len(datos) < 2:
        raise ValueError("Solo se obtuvo 1 fila válida; se requieren al menos 2 para covarianza.")

    return datos, filas_invalidas, total_entrada

def simple():
    """
    Versión rápida:
      - Pide ruta de archivo
      - NO usa encabezado
      - Usa TODAS las columnas como numéricas
      - Calcula la distancia de Mahalanobis de cada fila válida
    """
    print("=== Distancia de Mahalanobis (versión simple) ===")
    ruta = input("Ruta del archivo (.csv/.txt/.xlsx): ").strip()

    filas = cargar_filas_desde_archivo(ruta)

    max_cols = max(len(f) for f in filas)
    indices_columnas = list(range(max_cols))  # todas las columnas

    datos, filas_invalidas, total_entrada = construir_matriz_numerica(
        filas,
        usar_encabezado=False,
        indices_columnas=indices_columnas
    )

    print(f"\nFilas de entrada: {total_entrada}")
    print(f"Filas válidas usadas: {len(datos)}")
    print(f"Filas descartadas por datos no numéricos o incompletos: {filas_invalidas}")

    media = calcular_media(datos)
    cov = calcular_matriz_covarianza(datos, media)
    cov_inv = inversa_matriz(cov)

    print("\nDistancia de Mahalanobis de cada fila válida respecto a la media:")
    for i, fila in enumerate(datos):
        d = distancia_mahalanobis(fila, media, cov_inv)
        print(f"Fila {i+1}: {d:.6f}")

def completo():
    filas = None
    datos = None
    media = None
    cov_inv = None
    indices_columnas = None
    usar_encabezado = False

    while True:
        print("\n====================================")
        print("  MENÚ - Distancia de Mahalanobis")
        print("====================================")
        print("1) Cargar archivo (.csv / .txt / .xlsx)")
        print("2) Seleccionar columnas numéricas")
        print("3) Construir matriz numérica y calcular Σ⁻¹")
        print("4) Calcular distancia de Mahalanobis de todas las filas válidas")
        print("5) Mostrar resumen actual")
        print("0) Salir")
        opcion = input("Elige una opción: ").strip()

        if opcion == "0":
            print("Saliendo del programa.")
            break

        elif opcion == "1":
            ruta = input("Ruta del archivo: ").strip()
            try:
                filas = cargar_filas_desde_archivo(ruta)
                datos = None
                media = None
                cov_inv = None
                indices_columnas = None

                print(f"\nArchivo cargado correctamente.")
                print(f"Número total de filas: {len(filas)}")
                if filas:
                    print("Ejemplo de primera fila:")
                    print(filas[0])

                # Preguntar por encabezado
                while True:
                    resp = input("¿La primera fila es encabezado? (s/n): ").strip().lower()
                    if resp in ("s", "si", "sí"):
                        usar_encabezado = True
                        break
                    elif resp in ("n", "no"):
                        usar_encabezado = False
                        break
                    else:
                        print("Responde 's' o 'n'.")
            except Exception as e:
                print(f"\nERROR al cargar el archivo: {e}")

        elif opcion == "2":
            if filas is None:
                print("Primero debes cargar un archivo (opción 1).")
                continue
            try:
                indices_columnas = seleccionar_columnas(filas)
                print(f"Columnas seleccionadas (0-based): {indices_columnas}")
            except Exception as e:
                print(f"ERROR al seleccionar columnas: {e}")

        elif opcion == "3":
            if filas is None:
                print("Primero debes cargar un archivo (opción 1).")
                continue
            if indices_columnas is None:
                print("Primero debes seleccionar las columnas numéricas (opción 2).")
                continue
            try:
                datos, filas_invalidas, total_entrada = construir_matriz_numerica(
                    filas,
                    usar_encabezado,
                    indices_columnas
                )
                print("\nMatriz numérica construida.")
                print(f"Filas de entrada (sin contar encabezado): {total_entrada}")
                print(f"Filas válidas usadas: {len(datos)}")
                print(f"Filas descartadas: {filas_invalidas}")

                media = calcular_media(datos)
                cov = calcular_matriz_covarianza(datos, media)
                cov_inv = inversa_matriz(cov)

                print("Matriz de covarianza invertida Σ⁻¹ calculada correctamente.")
            except Exception as e:
                datos = None
                media = None
                cov_inv = None
                print(f"\nERROR al construir matriz o invertir Σ: {e}")

        elif opcion == "4":
            if datos is None or media is None or cov_inv is None:
                print("Primero debes construir la matriz numérica y Σ⁻¹ (opción 3).")
                continue
            try:
                print("\nDistancia de Mahalanobis de cada fila válida respecto a la media:")
                for i, fila in enumerate(datos):
                    try:
                        d = distancia_mahalanobis(fila, media, cov_inv)
                        print(f"Fila válida {i+1}: {d:.6f}")
                    except Exception as e:
                        print(f"  Error en fila {i+1}: {e}")
            except Exception as e:
                print(f"\nERROR al calcular distancias: {e}")

        elif opcion == "5":
            print("\n=== RESUMEN ACTUAL ===")
            if filas is None:
                print("Archivo: no cargado.")
            else:
                print(f"Archivo: cargado, filas totales: {len(filas)}")
                print("Encabezado:", "sí" if usar_encabezado else "no")
            if indices_columnas is None:
                print("Columnas numéricas: no seleccionadas.")
            else:
                print(f"Columnas numéricas (0-based): {indices_columnas}")
            if datos is None:
                print("Matriz numérica: no construida.")
            else:
                print(f"Matriz numérica: {len(datos)} filas válidas x {len(datos[0])} columnas.")
            if media is None:
                print("Media: no calculada.")
            else:
                print("Media calculada.")
            if cov_inv is None:
                print("Σ⁻¹: no calculada.")
            else:
                print("Σ⁻¹ calculada.")
            print("======================")

        else:
            print("Opción no válida. Intenta de nuevo.")

if __name__ == "__main__":
    # simple()
    completo()
