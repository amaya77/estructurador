"""Lectura, validación y exportación de archivos de Excel."""

import unicodedata
from dataclasses import dataclass
from itertools import chain, islice
from pathlib import Path
from tempfile import TemporaryDirectory


# Se valida su orden. Se permiten columnas adicionales después de estas cuatro.
ENCABEZADOS = (
    "userDocumentType",
    "userDocumentNumber",
    "companyDocumentType",
    "companyDocumentNumber",
)
ENCABEZADOS_SALIDA = (*ENCABEZADOS, "parameters")
REGISTROS_POR_ARCHIVO = 200_000
# MB decimales: margen por debajo del límite de carga de 2 MB.
PESO_MAXIMO_BYTES = 1_900_000


@dataclass(frozen=True)
class ArchivoGenerado:
    ruta: Path
    registros: int
    peso_bytes: int = 0


# Equivalencias para seguir leyendo las bases creadas con títulos en español.
ALIAS_ENCABEZADOS = {
    "tipo de documento usuario": "userdocumenttype",
    "documento usuario": "userdocumentnumber",
    "tipo de documento empresa": "companydocumenttype",
    "documento empresa": "companydocumentnumber",
}


@dataclass(frozen=True)
class ErrorValidacion:
    fila: int
    columna: int
    campo: str
    motivo: str
    permite_faltantes: bool = False


def texto(valor):
    """Convierte documentos numéricos a texto sin agregar un '.0'."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def normalizar_encabezado(valor):
    """Ignora mayúsculas, tildes y espacios repetidos al comparar títulos."""
    valor = unicodedata.normalize("NFD", texto(valor).casefold())
    sin_tildes = "".join(c for c in valor if not unicodedata.combining(c))
    normalizado = " ".join(sin_tildes.split())
    return ALIAS_ENCABEZADOS.get(normalizado, normalizado)


def leer_excel(ruta):
    """Lee la primera hoja. Las fórmulas se conservan para poder detectarlas."""
    ruta = Path(ruta)
    if ruta.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook

        libro = load_workbook(ruta, read_only=True, data_only=False)
        try:
            filas = list(libro.worksheets[0].iter_rows(values_only=True))
        finally:
            libro.close()
    elif ruta.suffix.lower() == ".xls":
        import xlrd

        libro = xlrd.open_workbook(str(ruta), on_demand=True)
        try:
            hoja = libro.sheet_by_index(0)
            filas = [hoja.row_values(i) for i in range(hoja.nrows)]
        finally:
            libro.release_resources()
    else:
        raise ValueError("Selecciona un archivo .xlsx o .xls.")

    # Excel puede guardar filas vacías al final por el formato de las celdas.
    # Las filas vacías intermedias sí se validan y se incluyen en el reporte.
    while filas and not any(texto(celda) for celda in filas[-1]):
        filas.pop()
    return filas


def cantidad_etiquetas(filas):
    """Cuenta desde E hasta la última columna con encabezado o datos.

    Ignora las columnas vacías del final que Excel guarda por su formato.
    Incluye los huecos intermedios para detectar encabezados faltantes.
    """
    ultima_columna = 4
    for fila in filas:
        for indice in range(len(fila) - 1, 3, -1):
            if texto(fila[indice]):
                ultima_columna = max(ultima_columna, indice + 1)
                break
    return ultima_columna - 4


def iterar_consolidados(filas):
    """Consolida una fila a la vez para reducir el uso de memoria."""
    cantidad = cantidad_etiquetas(filas)
    claves = [texto(filas[0][i] if i < len(filas[0]) else None) for i in range(4, 4 + cantidad)]
    for fila in islice(filas, 1, None):
        campos = []
        for indice, clave in enumerate(claves, start=4):
            if clave:
                valor = texto(fila[indice] if indice < len(fila) else None)
                campos.append(f"{clave}:{valor}")
        yield [*(texto(valor) for valor in fila[:4]), ";".join(campos)]


def consolidar_etiquetas(filas):
    """Renombra A-E y reúne las etiquetas en E como clave:valor;clave:valor."""
    return [list(ENCABEZADOS_SALIDA), *iterar_consolidados(filas)]


def validar(filas):
    """Devuelve todos los errores, usando filas y columnas de Excel (desde 1)."""
    errores = []
    if not filas:
        return [ErrorValidacion(1, 1, "Archivo", "El archivo está vacío.")]

    encabezados = filas[0]
    # Buscar cada etiqueta en toda la fila permite distinguir un título ausente
    # de uno que sí existe, pero fue ubicado en otra columna.
    normalizados = [normalizar_encabezado(valor) for valor in encabezados]
    for indice, esperado in enumerate(ENCABEZADOS):
        recibido = encabezados[indice] if indice < len(encabezados) else None
        etiqueta = normalizar_encabezado(esperado)
        if normalizar_encabezado(recibido) != etiqueta:
            posiciones = [str(i + 1) for i, titulo in enumerate(normalizados) if titulo == etiqueta]
            if posiciones:
                motivo = (
                    f"Orden incorrecto: '{esperado}' debe estar en la columna {indice + 1}; "
                    f"se encontró en la(s) columna(s) {', '.join(posiciones)}. "
                    f"En la columna {indice + 1} aparece '{texto(recibido) or '(vacía)'}'. "
                    "Reordena las columnas completas, incluidos sus datos."
                )
            else:
                motivo = (
                    f"Estructura incorrecta: falta la etiqueta '{esperado}', "
                    f"esperada en la columna {indice + 1}. "
                    f"Se encontró '{texto(recibido) or '(vacía)'}'."
                )
            errores.append(ErrorValidacion(
                1, indice + 1, esperado, motivo,
            ))

    # Sin títulos correctos no podemos atribuir los datos al campo adecuado.
    if errores:
        return errores
    if len(filas) == 1:
        return [ErrorValidacion(2, 1, "Archivo", "No hay registros para convertir.")]

    cantidad = cantidad_etiquetas(filas)
    for indice in range(4, 4 + cantidad):
        clave = texto(encabezados[indice] if indice < len(encabezados) else None)
        if not clave:
            errores.append(ErrorValidacion(
                1, indice + 1, "Etiqueta",
                "Falta el encabezado de la etiqueta. Cada columna desde E debe tener una clave.",
                permite_faltantes=True,
            ))

    for numero_fila, fila in enumerate(islice(filas, 1, None), start=2):
        for indice, campo in enumerate(ENCABEZADOS):
            valor = texto(fila[indice] if indice < len(fila) else None)
            if not valor:
                motivo = "La celda está vacía."
            elif valor.startswith("="):
                motivo = "Usa un valor fijo en esta celda, no una fórmula."
            elif indice in (1, 3) and len(valor) > 20:
                motivo = f"El documento tiene {len(valor)} caracteres; el máximo es 20."
            else:
                # El tipo de documento es libre: no hay catálogo de valores.
                continue
            errores.append(ErrorValidacion(numero_fila, indice + 1, campo, motivo))
        for indice in range(4, 4 + cantidad):
            valor = texto(fila[indice] if indice < len(fila) else None)
            if not valor:
                clave = texto(encabezados[indice] if indice < len(encabezados) else None)
                errores.append(ErrorValidacion(
                    numero_fila, indice + 1, clave or "Etiqueta sin encabezado",
                    "El valor del campo personalizado está vacío. Completa la celda para generar el archivo.",
                    permite_faltantes=True,
                ))
    return errores


def guardar_xlsx(ruta, filas):
    """Escribe secuencialmente un XLSX con valores de texto y sin fórmulas."""
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell

    libro = Workbook(write_only=True)
    hoja = libro.create_sheet("Estructurado")
    for fila in filas:
        celdas = []
        for valor in fila:
            celda = WriteOnlyCell(hoja, value=texto(valor))
            # Conserva ceros, documentos largos y textos que comienzan con '='.
            celda.data_type = "s"
            celda.number_format = "@"
            celdas.append(celda)
        hoja.append(celdas)
    libro.save(ruta)
    libro.close()


def rutas_particiones(ruta, cantidad):
    """Elige un nombre de lote libre sin sobrescribir salidas anteriores."""
    intento = 1
    while True:
        sufijo = "" if intento == 1 else f" ({intento})"
        rutas = [ruta.with_name(f"{ruta.stem} estructurado{i}{sufijo}.xlsx")
                 for i in range(1, cantidad + 1)]
        if not any(candidata.exists() for candidata in rutas):
            return rutas
        intento += 1


def convertir(ruta):
    """Retorna (lista de ArchivoGenerado, errores). Todo error bloquea la salida."""
    ruta = Path(ruta)
    filas = leer_excel(ruta)
    errores = validar(filas)
    if errores:
        return [], errores

    # Validar toda la base antes de guardar cualquier partición. No se permite
    # que openpyxl recorte silenciosamente una celda demasiado larga.
    for i, fila in enumerate(iterar_consolidados(filas), start=2):
        for j, valor in enumerate(fila, start=1):
            if len(valor) > 32767:
                errores.append(ErrorValidacion(
                    i, j, "parameters" if j == 5 else "Celda",
                    "El formato .xlsx admite hasta 32.767 caracteres por celda. El contenido resultante supera ese límite.",
                ))
    if errores:
        return [], errores

    total = len(filas) - 1
    registros = iterar_consolidados(filas)
    # Preparar el lote completo antes de colocar los archivos junto al origen.
    with TemporaryDirectory(prefix=".estructurador-", dir=ruta.parent) as carpeta:
        temporales = []

        def guardar_parte(bloque, inicio, fin, fila_inicial):
            """Mide el XLSX cerrado y subdivide los bloques que excedan el peso."""
            temporal = Path(carpeta) / f"parte{len(temporales) + 1}.xlsx"
            guardar_xlsx(temporal, chain(
                [ENCABEZADOS_SALIDA], islice(bloque, inicio, fin),
            ))
            peso = temporal.stat().st_size
            if peso <= PESO_MAXIMO_BYTES:
                temporales.append((temporal, fin - inicio, peso))
                return True
            if fin - inicio == 1:
                errores.append(ErrorValidacion(
                    fila_inicial + inicio, 1, "Archivo",
                    "Este registro, junto con el encabezado, supera el peso máximo "
                    f"de {PESO_MAXIMO_BYTES:,} bytes por archivo. Reduce su contenido.",
                ))
                return False
            medio = (inicio + fin) // 2
            return (guardar_parte(bloque, inicio, medio, fila_inicial)
                    and guardar_parte(bloque, medio, fin, fila_inicial))

        for inicio in range(0, total, REGISTROS_POR_ARCHIVO):
            bloque = list(islice(registros, REGISTROS_POR_ARCHIVO))
            if not guardar_parte(bloque, 0, len(bloque), inicio + 2):
                return [], errores
        rutas = rutas_particiones(ruta, len(temporales))
        salidas = []
        for (temporal, cantidad, peso), destino in zip(temporales, rutas):
            temporal.rename(destino)
            salidas.append(ArchivoGenerado(destino, cantidad, peso))
    return salidas, []
