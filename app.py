"""Interfaz de escritorio. Ejecutar con: python app.py"""

import queue
from itertools import islice
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

from PIL import Image, ImageTk

from validador import contar_registros, convertir


def leer_vista_previa(ruta):
    """Lee solo el encabezado y los primeros diez registros de una partición."""
    from openpyxl import load_workbook

    libro = load_workbook(ruta, read_only=True, data_only=False)
    try:
        return list(islice(libro.worksheets[0].iter_rows(values_only=True), 11))
    finally:
        libro.close()


class Aplicacion:
    def __init__(self, ventana):
        self.ventana = ventana
        self.ruta = None
        # El trabajo ocurre en otro hilo para no congelar la ventana.
        self.resultados = queue.Queue()
        self.avances = queue.Queue()
        ventana.title("Estructurador centro de mensajes")
        self.icono = tk.PhotoImage(file=str(Path(__file__).resolve().parent / "assets" / "bancolombia.png"))
        ventana.iconphoto(True, self.icono)
        ventana.geometry("1000x710")
        ventana.minsize(720, 570)
        self.fondo_original = Image.open(Path(__file__).resolve().parent / "FONDO.png").convert("RGB")
        self.redimension_pendiente = None
        marco = tk.Canvas(ventana, background="#202020", highlightthickness=0)
        self.fondo_canvas = marco
        self.fondo_elemento = marco.create_image(0, 0, anchor="nw")
        marco.bind("<Configure>", self.programar_fondo)
        marco.pack(fill="both", expand=True)
        cabecera = tk.Frame(marco, background="#202020")
        cabecera.pack(fill="x", padx=24, pady=(24, 0))
        tk.Label(cabecera, text="Estructurador centro de mensajes", font=("Segoe UI", 17, "bold"), background="#202020", foreground="white").pack(side="left")
        self.nueva = ttk.Button(cabecera, text="Crear nueva estructura", command=self.reiniciar)
        self.nueva.pack(side="right", padx=(12, 0))
        tk.Label(marco, text="Gestionar la información de tus bases nunca había sido tan fácil", background="#202020", foreground="white").pack(anchor="w", padx=24, pady=(8, 16))
        self.archivo = tk.StringVar(value="Ningún archivo seleccionado")
        tk.Label(marco, textvariable=self.archivo, wraplength=570, justify="left", background="#202020", foreground="white").pack(anchor="w", padx=24)
        botones = tk.Frame(marco, background="#202020")
        botones.pack(anchor="w", padx=24, pady=18)
        self.buscar = ttk.Button(botones, text="Cargar archivo", command=self.seleccionar)
        self.buscar.pack(side="left", padx=(0, 12))
        self.procesar = ttk.Button(botones, text="Convertir", command=self.iniciar, state="disabled")
        self.procesar.pack(side="left")
        self.estado = tk.StringVar(value="Los resultados se guardan junto al archivo original.")
        tk.Label(marco, textvariable=self.estado, wraplength=850, justify="left", font=("Segoe UI", 10, "bold"), background="#202020", foreground="white").pack(anchor="w", padx=24)

        self.texto_progreso = tk.StringVar(value="")
        tk.Label(marco, textvariable=self.texto_progreso, background="#202020",
                 foreground="white", anchor="w").pack(fill="x", padx=24, pady=(8, 2))
        self.barra_progreso = ttk.Progressbar(marco, maximum=100, mode="determinate")
        self.barra_progreso.pack(fill="x", padx=(24, 64), pady=(0, 4))

        # Los errores se muestran en pantalla, sin generar un CSV.
        # Las pestañas permanecen ocultas hasta que hay un resultado que mostrar.
        self.panel_resultados = ttk.Notebook(marco)
        self.monitor, self.resumen, self.tabla_errores = self.crear_tabla_incidencias()
        self.monitor_faltantes, self.resumen_faltantes, self.tabla_faltantes = self.crear_tabla_incidencias()

        self.previsualizacion = ttk.Frame(self.panel_resultados, padding=10)
        self.resumen_previa = tk.StringVar()
        ttk.Label(self.previsualizacion, textvariable=self.resumen_previa, wraplength=850).pack(anchor="w", pady=(0, 8))
        contenedor = ttk.Frame(self.previsualizacion)
        contenedor.pack(fill="both", expand=True)
        contenedor.rowconfigure(0, weight=1)
        contenedor.columnconfigure(0, weight=1)
        self.tabla_previa = ttk.Treeview(contenedor, show="headings", height=8)
        vertical_previa = ttk.Scrollbar(contenedor, orient="vertical", command=self.tabla_previa.yview)
        horizontal_previa = ttk.Scrollbar(contenedor, orient="horizontal", command=self.tabla_previa.xview)
        self.tabla_previa.configure(yscrollcommand=vertical_previa.set, xscrollcommand=horizontal_previa.set)
        self.tabla_previa.grid(row=0, column=0, sticky="nsew")
        vertical_previa.grid(row=0, column=1, sticky="ns")
        horizontal_previa.grid(row=1, column=0, sticky="ew")
        self.activar_scroll(self.tabla_previa)

    def mostrar_progreso(self, porcentaje, mensaje):
        """La lectura indica actividad; las otras etapas informan avance real."""
        if porcentaje is None:
            if str(self.barra_progreso["mode"]) != "indeterminate":
                self.barra_progreso.configure(mode="indeterminate", value=0)
                self.barra_progreso.start(12)
            self.texto_progreso.set(mensaje)
        else:
            self.barra_progreso.stop()
            self.barra_progreso.configure(mode="determinate", value=porcentaje)
            self.texto_progreso.set(f"{mensaje} ({porcentaje:.0f} % de esta etapa)" if mensaje else "")

    def programar_fondo(self, evento):
        """Agrupa los cambios de tamaño para redibujar sin bloquear el arrastre."""
        if self.redimension_pendiente is not None:
            self.ventana.after_cancel(self.redimension_pendiente)
        self.redimension_pendiente = self.ventana.after(
            60, self.actualizar_fondo, evento.width, evento.height,
        )

    def actualizar_fondo(self, ancho, alto):
        """Ajusta la imagen a la ventana y conserva una referencia para Tk."""
        self.redimension_pendiente = None
        imagen = self.fondo_original.resize((max(1, ancho), max(1, alto)), Image.Resampling.LANCZOS)
        self.fondo_imagen = ImageTk.PhotoImage(imagen, master=self.ventana)
        self.fondo_canvas.itemconfigure(self.fondo_elemento, image=self.fondo_imagen)
        self.fondo_canvas.tag_lower(self.fondo_elemento)

    def activar_scroll(self, tabla):
        """La rueda desplaza filas; Mayús + rueda desplaza horizontalmente."""
        def desplazar(evento, horizontal=False):
            pasos = -int(evento.delta / 120)
            if not pasos and evento.delta:
                pasos = -1 if evento.delta > 0 else 1
            if horizontal:
                tabla.xview_scroll(pasos, "units")
            else:
                tabla.yview_scroll(pasos, "units")
            return "break"

        tabla.bind("<MouseWheel>", desplazar)
        tabla.bind("<Shift-MouseWheel>", lambda evento: desplazar(evento, horizontal=True))

    def ajustar_anchos(self, tabla):
        """Permite llegar al final de textos largos con la barra horizontal."""
        fuente = tkfont.nametofont("TkDefaultFont")
        columnas = tabla["columns"]
        anchos = [max(80, fuente.measure(tabla.heading(clave, "text")) + 24) for clave in columnas]
        for registro in tabla.get_children():
            for indice, valor in enumerate(tabla.item(registro, "values")):
                # Treeview muestra una línea por fila: medir la más larga.
                ancho = max((fuente.measure(linea) for linea in str(valor).splitlines()), default=0)
                anchos[indice] = max(anchos[indice], ancho + 24)
        for clave, ancho in zip(columnas, anchos):
            tabla.column(clave, width=ancho, minwidth=80, stretch=False)

    def crear_tabla_incidencias(self):
        """Crea una tabla independiente para errores o para etiquetas faltantes."""
        monitor = ttk.Frame(self.panel_resultados, padding=10)
        resumen = tk.StringVar()
        ttk.Label(monitor, textvariable=resumen, wraplength=650).pack(anchor="w", pady=(0, 8))
        tabla = ttk.Frame(monitor)
        tabla.pack(fill="both", expand=True)
        tabla.rowconfigure(0, weight=1)
        tabla.columnconfigure(0, weight=1)
        listado = ttk.Treeview(
            tabla, columns=("fila", "columna", "campo", "motivo"),
            show="headings", selectmode="browse", height=8,
        )
        for clave, titulo, ancho in (
            ("fila", "Fila", 65), ("columna", "Columna", 75),
            ("campo", "Campo", 210), ("motivo", "Motivo del error", 550),
        ):
            listado.heading(clave, text=titulo)
            listado.column(clave, width=ancho, minwidth=ancho, stretch=clave == "motivo")
        vertical = ttk.Scrollbar(tabla, orient="vertical", command=listado.yview)
        horizontal = ttk.Scrollbar(tabla, orient="horizontal", command=listado.xview)
        listado.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        listado.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.activar_scroll(listado)

        return monitor, resumen, listado

    def reiniciar(self):
        """Vuelve al estado inicial sin borrar archivos que ya se generaron."""
        self.ruta = None
        self.mostrar_progreso(0, "")
        self.archivo.set("Ningún archivo seleccionado")
        self.procesar.config(state="disabled")
        self.estado.set("Los resultados se guardan junto al archivo original.")
        self.limpiar_reporte("")

    def seleccionar(self):
        ruta = filedialog.askopenfilename(title="Selecciona un Excel", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            seleccion = Path(ruta)
            try:
                peso = seleccion.stat().st_size
            except OSError as error:
                messagebox.showerror("Error", f"No se pudo acceder al archivo seleccionado.\n\nDetalle: {error}")
                return
            self.ruta = seleccion
            self.archivo.set(f"{self.ruta}\nPeso del archivo: {peso / 1_000_000:.2f} MB ({peso:,} bytes)\nRegistros: contando…")
            self.procesar.config(state="disabled")
            self.buscar.config(state="disabled")
            self.nueva.config(state="disabled")
            self.estado.set("Contando registros de la primera hoja…")
            self.mostrar_progreso(None, "Leyendo y contando registros…")
            self.limpiar_reporte("")
            resultados = queue.Queue()
            threading.Thread(target=self.contar_archivo, args=(seleccion, resultados), daemon=True).start()
            self.ventana.after(100, self.recibir_conteo, seleccion, peso, resultados)

    @staticmethod
    def contar_archivo(ruta, resultados):
        try:
            resultados.put((True, contar_registros(ruta)))
        except Exception as error:
            resultados.put((False, str(error)))

    def recibir_conteo(self, ruta, peso, resultados):
        """Actualiza los controles solo desde el hilo principal de Tk."""
        try:
            correcto, resultado = resultados.get_nowait()
        except queue.Empty:
            self.ventana.after(100, self.recibir_conteo, ruta, peso, resultados)
            return
        self.buscar.config(state="normal")
        self.nueva.config(state="normal")
        self.mostrar_progreso(100 if correcto else 0, "Conteo finalizado" if correcto else "Lectura interrumpida")
        detalle = f"{ruta}\nPeso del archivo: {peso / 1_000_000:.2f} MB ({peso:,} bytes)"
        if not correcto:
            self.archivo.set(f"{detalle}\nRegistros: no disponibles")
            self.estado.set("No se pudo leer el archivo. Selecciona un Excel válido.")
            messagebox.showerror("Error", f"No se pudo contar los registros.\n\nDetalle: {resultado}")
            return
        self.archivo.set(f"{detalle}\nRegistros: {resultado:,} (sin encabezado)")
        self.procesar.config(state="normal")
        self.estado.set("Listo para validar y convertir.")
        self.resumen.set("Pulsa Convertir para validar el archivo seleccionado.")

    def limpiar_reporte(self, mensaje):
        """Limpia y oculta los resultados anteriores al comenzar otro flujo."""
        for tabla in (self.tabla_errores, self.tabla_faltantes, self.tabla_previa):
            registros = tabla.get_children()
            if registros:
                tabla.delete(*registros)
            tabla.xview_moveto(0)
            tabla.yview_moveto(0)
        for pestaña in self.panel_resultados.tabs():
            self.panel_resultados.forget(pestaña)
        self.panel_resultados.pack_forget()
        self.resumen_previa.set("")
        self.resumen_faltantes.set("")
        self.resumen.set(mensaje)

    def iniciar(self):
        self.nueva.config(state="disabled")
        self.buscar.config(state="disabled")
        self.procesar.config(state="disabled")
        self.mostrar_progreso(None, "Leyendo archivo Excel…")
        self.avances = queue.Queue()
        self.estado.set("Validando archivo…")
        self.limpiar_reporte("Validando archivo…")
        threading.Thread(
            target=self.trabajar, daemon=True,
        ).start()
        self.ventana.after(100, self.recibir_resultado)

    def trabajar(self):
        try:
            salidas, errores = convertir(self.ruta, progreso=lambda valor, mensaje: self.avances.put((valor, mensaje)))
            if salidas:
                self.avances.put((None, "Preparando vista previa…"))
            previa = leer_vista_previa(salidas[0].ruta) if salidas else None
            self.resultados.put((True, (salidas, errores, previa)))
        except Exception as error:
            self.resultados.put((False, str(error)))

    def recibir_resultado(self):
        ultimo = None
        while True:
            try:
                ultimo = self.avances.get_nowait()
            except queue.Empty:
                break
        if ultimo is not None:
            self.mostrar_progreso(*ultimo)
        try:
            correcto, resultado = self.resultados.get_nowait()
        except queue.Empty:
            self.ventana.after(100, self.recibir_resultado)
            return
        self.buscar.config(state="normal")
        self.nueva.config(state="normal")
        self.procesar.config(state="normal")
        if not correcto:
            self.mostrar_progreso(0, "Proceso interrumpido")
            self.estado.set("No se pudo procesar el archivo.")
            self.resumen.set("No fue posible completar la validación.")
            self.tabla_errores.insert("", "end", values=("—", "—", "Archivo", resultado))
            self.panel_resultados.add(self.monitor, text="Reporte de errores")
            self.panel_resultados.pack(fill="both", expand=True, padx=(24, 64), pady=(16, 56))
            messagebox.showerror("Error", f"Revisa que el archivo sea válido y que tengas permiso para guardar en su carpeta.\n\nDetalle: {resultado}")
            return
        salidas, errores, previa = resultado
        self.mostrar_progreso(100 if salidas else 0, "Proceso completado" if salidas else "Corrige los errores para continuar")
        if salidas:
            total = sum(salida.registros for salida in salidas)
            self.estado.set(
                f"Se generaron {len(salidas)} archivo(s) con {total:,} registros en total. "
                f"Carpeta: {salidas[0].ruta.parent}"
            )
        else:
            self.estado.set("No se generó el archivo estructurado. Corrige los errores del reporte en pantalla.")
        if previa:
            filas = previa
            claves = [f"columna_{i}" for i in range(len(filas[0]))]
            self.tabla_previa.configure(columns=claves)
            for indice, (clave, titulo) in enumerate(zip(claves, filas[0])):
                self.tabla_previa.heading(clave, text=titulo)
                self.tabla_previa.column(clave, width=520 if indice == 4 else 190, minwidth=150)
            for fila in filas[1:]:
                self.tabla_previa.insert("", "end", values=fila)
            self.ajustar_anchos(self.tabla_previa)
            detalle = "\n".join(
                f"{salida.ruta.name}: {salida.registros:,} registros · {salida.peso_bytes / 1_000_000:.2f} MB"
                for salida in salidas
            )
            self.resumen_previa.set(
                f"{detalle}\nReparto estimado por peso del original; el peso final de las partes puede variar."
                f"\nVista previa de {salidas[0].ruta.name}: primeros {len(filas) - 1} registros."
            )
            self.panel_resultados.add(self.previsualizacion, text="Vista previa del archivo")
        # Los faltantes de etiquetas tienen su propia pestaña, independiente
        # de los errores de documentos, estructura o límites del archivo.
        errores_generales = [error for error in errores if not error.permite_faltantes]
        faltantes = [error for error in errores if error.permite_faltantes]
        for incidencias, monitor, tabla, resumen, titulo in (
            (errores_generales, self.monitor, self.tabla_errores, self.resumen, "Errores"),
            (faltantes, self.monitor_faltantes, self.tabla_faltantes, self.resumen_faltantes, "Etiquetas faltantes"),
        ):
            if not incidencias:
                continue
            self.panel_resultados.add(monitor, text=f"{titulo} ({len(incidencias)})")
            for error in incidencias:
                tabla.insert("", "end", values=(
                    error.fila, error.columna, error.campo, error.motivo,
                ))
            self.ajustar_anchos(tabla)
            if titulo == "Errores":
                resumen.set(f"Se encontraron {len(incidencias)} errores que impiden generar el archivo.")
            else:
                resumen.set(f"Se encontraron {len(incidencias)} faltantes de etiquetas. No se generó el archivo. Completa los encabezados y valores vacíos y vuelve a convertir.")
        if self.panel_resultados.tabs():
            self.panel_resultados.pack(fill="both", expand=True, padx=(24, 64), pady=(16, 56))
            self.panel_resultados.select(self.panel_resultados.tabs()[0])


if __name__ == "__main__":
    ventana = tk.Tk()
    Aplicacion(ventana)
    ventana.mainloop()
