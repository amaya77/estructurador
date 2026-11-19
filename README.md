# Estructurador de Excel

Para generar un ejecutable portable de Windows, consulta `LEEME_WINDOWS.md` y
ejecuta `CREAR_EXE_WINDOWS.bat` en Windows. El resultado sera
`dist/Estructurador.exe`, con el fondo y el icono incluidos.

En macOS con Python 3.14 de Homebrew, prepara el entorno así:

```sh
brew install python-tk@3.14
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
```

Si VS Code ya tenía otro intérprete seleccionado, ejecuta **Python: Select
Interpreter** y elige `.venv/bin/python`. El componente Tk se instala con
Homebrew; no se instala mediante `pip`.

Aplicación de escritorio en Python. Instala las dependencias y abre la ventana:

```powershell
python -m pip install -r requirements.txt
python app.py
```

Selecciona un archivo `.xlsx` o `.xls` y pulsa **Convertir**.
Al seleccionarlo, la ventana muestra su peso en MB y en bytes junto a la ruta.
Este es el peso del archivo de origen; las partes generadas se miden por separado.
El archivo debe tener
los títulos en la fila 1 de la primera hoja, en este orden:

1. userDocumentType (también acepta Tipo de documento usuario)
2. userDocumentNumber (también acepta Documento usuario)
3. companyDocumentType (también acepta Tipo de documento empresa)
4. companyDocumentNumber (también acepta Documento empresa)

La salida siempre tiene estos cinco encabezados exactos: `userDocumentType`,
`userDocumentNumber`, `companyDocumentType`, `companyDocumentNumber`, `parameters`.

Los títulos pueden variar en mayúsculas, tildes y espacios repetidos. Se permiten
columnas adicionales. Los cuatro campos son obligatorios; ambos documentos
admiten hasta 20 caracteres. El tipo de documento admite cualquier valor no vacío.
Las fórmulas en esos cuatro campos deben reemplazarse por valores fijos.

La interfaz no tiene casillas de reporte ni de generación con faltantes.
Los errores se muestran únicamente en pantalla; no se genera un CSV.
Los encabezados y valores vacíos de los campos personalizados bloquean la
generación, igual que los errores de documentos y los límites de Excel.
Al cargar otro archivo o volver a convertir, se limpia el reporte anterior.
Si la estructura es incorrecta, primero se reportan
los encabezados; al corregirlos se podrán validar los registros. Si todo es válido,
se generan archivos `.xlsx` de **máximo 1,9 MB (1.900.000 bytes)**, incluido
el encabezado y todos los componentes del archivo. La partición depende solo
del peso, sin un límite de cantidad de registros. El tamaño se comprueba sobre cada XLSX terminado;
si supera el máximo, el bloque se divide en mitades y se vuelven a guardar y medir
las partes hasta que todas cumplan el límite de peso.

Los nombres siguen siendo `base estructurado1.xlsx`, `base estructurado2.xlsx`,
etc. La cantidad de archivos depende del peso de los datos y su compresión:
una parte puede contener más de 200.000 registros si no supera 1,9 MB.
Se conservan el orden, los documentos como texto y el encabezado en cada archivo.
Si un registro individual no cabe junto al encabezado, se informa su fila y no
se publica ninguna parte del lote.

Las partes se guardan junto al original. Si ya existe una parte con el mismo
nombre, el lote nuevo añade un sufijo común, como `base estructurado1 (2).xlsx`.
Se valida toda la base antes de escribir las partes: un error en cualquier
registro bloquea el lote completo. La salida se mantiene en formato `.xlsx`.

La ventana muestra los nombres, cantidades y pesos en MB de todas las partes y una vista
previa de los primeros diez registros de la primera parte, con las cinco
columnas resultantes. Los campos personalizados vacíos aparecen en la pestaña
**Etiquetas faltantes** y bloquean la generación. Los problemas de documentos, estructura o
límites aparecen en **Errores**. Cada pestaña muestra su cantidad y aparece solo
si tiene incidencias; si hay ambos tipos, se muestran ambas. Al iniciar no se muestran tablas vacías.
El botón superior **Crear nueva estructura** limpia la selección y los resultados, y permite empezar de cero sin borrar archivos generados.
Durante una conversión el botón permanece desactivado hasta que termine.

Cuando las etiquetas están fuera de orden, el reporte indica la columna esperada,
la ubicación encontrada y el título que ocupa esa posición. Reordena las columnas
completas junto con sus datos; cambiar solo los títulos podría asignar documentos
al campo equivocado. Las etiquetas ausentes se reportan como errores de estructura.

Las columnas desde E contienen etiquetas personalizables. Se cuentan hasta la
última columna con encabezado o datos, ignorando las columnas vacías del final.
Cada encabezado es una clave y cada celda es su valor. La salida conserva A-D y
reemplaza todas las columnas adicionales por una columna E llamada **parameters**.
Por ejemplo, las columnas `nombre`, `primer apellido` y `segundo apellido` producen:

```text
nombre:Ana;primer apellido:Pérez;segundo apellido:Gómez
```

Se conserva el orden de las etiquetas, las mayúsculas y los acentos; se eliminan
espacios al principio y al final de claves y valores. Los valores vacíos bloquean la generación. No se agregan llaves al
principio ni al final. Si no hay etiquetas, E queda vacía. Un encabezado vacío dentro del rango de etiquetas se reporta como
error. El texto se arma literalmente,
sin comillas ni reglas de escape para los caracteres `:`, `;`, `{` y `}`.
Los pares se separan por `;` sin agregar espacios; se conservan los espacios
internos de claves y valores como `primer apellido` o `Ana María`.
Convierte siempre el archivo de origen: volver a convertir el resultado trataría
su columna `parameters` como una nueva clave.

Las tres tablas (Errores, Etiquetas faltantes y Vista previa del archivo) tienen
barras vertical y horizontal. La rueda desplaza las filas y **Mayús + rueda**
desplaza las columnas. Los anchos se ajustan al contenido para poder consultar
todo el texto con la barra horizontal, incluidos los consolidados largos.

Solo se exportan los valores de la primera hoja. No se conservan estilos, macros
ni fórmulas ejecutables. Las fórmulas fuera de los cuatro campos se incorporan
como texto. Las fechas numéricas de archivos `.xls` se conservan como sus valores
internos. La salida `.xlsx` admite hasta 32.767 caracteres por celda;
se reporta un error si el texto consolidado de una fila excede ese límite.

Guarda los documentos como **Texto** en el Excel de origen: los ceros iniciales y
los dígitos que Excel ya haya eliminado o redondeado no se pueden recuperar.
Los espacios al comienzo y al final de un documento se eliminan al convertir.
Las filas totalmente vacías al final se ignoran; las intermedias se reportan.

El código está comentado: `app.py` contiene la interfaz y `validador.py` contiene
las reglas y la lectura/escritura. El original nunca se modifica.

La ventana usa `FONDO.png`, ubicado junto a `app.py`, como imagen de fondo.
La imagen se ajusta al tamaño de la ventana; los textos se muestran en blanco
y las tablas conservan su propio fondo para facilitar la lectura. Mantén
`FONDO.png` y la carpeta `assets` al distribuir la aplicación.
