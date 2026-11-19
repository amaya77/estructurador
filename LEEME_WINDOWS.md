# Crear Estructurador.exe

Estos pasos se realizan en un equipo con Windows de 64 bits.

1. Instala Python 3.14 de 64 bits desde https://www.python.org/downloads/windows/.
   Incluye el lanzador de Python y el componente Tcl/Tk de la instalacion.
2. Extrae todo el ZIP en una carpeta de tu equipo.
3. Haz doble clic en `CREAR_EXE_WINDOWS.bat`. Necesita internet para descargar
   las dependencias. Ejecuta las pruebas y compila el programa.
4. Espera el mensaje `Listo: dist\Estructurador.exe`.

Para usarlo o compartirlo, copia solamente `dist\Estructurador.exe`.
El equipo que lo utilice no necesita Python, un editor ni una instalacion del
programa. El fondo y el icono estan incluidos dentro del ejecutable.

Los resultados se guardan junto al Excel seleccionado, como en la version Python.
La salida sigue siendo `.xlsx`. La cantidad de archivos se calcula dividiendo
el peso del original entre 1.900.000 bytes y redondeando hacia arriba. Los registros
se reparten equilibradamente, sin limite de 200.000 registros. Cada parte se
escribe una sola vez. El peso final se muestra en pantalla y puede superar
1,9 MB: es una estimacion, no un tope garantizado. Los errores, incluidos los
campos personalizados vacios, bloquean la generacion y se muestran en pantalla.

Comprobacion antes de distribuir: abre el EXE, comprueba el fondo y el icono,
convierte una base valida y revisa la salida. Prueba tambien una base con campos
personalizados vacios para comprobar el reporte. La compilacion y el uso final
deben verificarse en Windows; el ZIP contiene los archivos de compilacion,
no un ejecutable ya compilado.

Para incorporar este cambio a un ejecutable anterior, vuelve a ejecutar
`CREAR_EXE_WINDOWS.bat` con estos archivos actualizados y distribuye el nuevo
`dist\Estructurador.exe`. El EXE anterior no se actualiza automaticamente.
