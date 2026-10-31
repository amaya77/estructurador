"""Pruebas de las reglas y de la conversión real, sin archivos del usuario."""

import itertools
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path

from openpyxl import Workbook, load_workbook

from validador import ENCABEZADOS, cantidad_etiquetas, consolidar_etiquetas, convertir, leer_excel, validar


class ValidacionTests(unittest.TestCase):
    def test_tipo_libre_y_limite_inclusivo(self):
        self.assertEqual(validar([ENCABEZADOS, ["Otro tipo", "A" * 20, "Libre", "000123"]]), [])

    def test_ambos_documentos_largos_y_campos_vacios(self):
        errores = validar([ENCABEZADOS, ["", "A" * 21, None, "B" * 21]])
        self.assertEqual([(e.fila, e.columna) for e in errores], [(2, 1), (2, 2), (2, 3), (2, 4)])

    def test_orden_incorrecto_y_columnas_faltantes(self):
        self.assertEqual(len(validar([[ENCABEZADOS[1], ENCABEZADOS[0]]])), 4)

    def test_todas_las_permutaciones_de_encabezados(self):
        for titulos in itertools.permutations(ENCABEZADOS):
            with self.subTest(titulos=titulos):
                errores = validar([titulos, ["CC", "123", "NIT", "456"]])
                fuera_de_orden = [i for i in range(4) if titulos[i] != ENCABEZADOS[i]]
                self.assertEqual([e.columna for e in errores], [i + 1 for i in fuera_de_orden])
                for error in errores:
                    self.assertEqual(error.fila, 1)
                    self.assertIn("Orden incorrecto", error.motivo)
                    self.assertIn(f"debe estar en la columna {error.columna}", error.motivo)
                    self.assertIn(f"columna(s) {titulos.index(error.campo) + 1}", error.motivo)

    def test_encabezado_ausente_no_se_confunde_con_desorden(self):
        errores = validar([["Nombre incorrecto", *ENCABEZADOS[1:]]])
        self.assertEqual(len(errores), 1)
        self.assertIn("falta la etiqueta 'userDocumentType'", errores[0].motivo)


    def test_normalizacion_y_columnas_adicionales(self):
        titulos = ["  TIPO  DE DOCUMENTO USUARIO ", "DOCUMENTO USUARIO", "Tipo de documénto empresa", "Documento empresa", "Extra"]
        self.assertEqual(validar([titulos, ["CC", "123", "NIT", "456", "Dato"]]), [])

    def test_archivo_vacio_sin_registros_y_fila_vacia_intermedia(self):
        self.assertTrue(validar([]))
        self.assertTrue(validar([ENCABEZADOS]))
        errores = validar([ENCABEZADOS, [None] * 4, ["CC", "1", "NIT", "2"]])
        self.assertEqual(len(errores), 4)
        self.assertTrue(all(e.fila == 2 for e in errores))

    def test_formula_requiere_valor_fijo(self):
        errores = validar([ENCABEZADOS, ["CC", "=1+1", "NIT", "2"]])
        self.assertEqual(errores[0].columna, 2)


    def test_consolidacion_por_fila_y_conteo(self):
        filas = [
            [*ENCABEZADOS, "nombre", "primer apellido", "segundo apellido", None],
            ["CC", "0001", "NIT", "123", "Ana", "Pérez", "Gómez", None],
            ["CC", "0002", "NIT", "456", "Luis", "Díaz", None, None],
        ]
        self.assertEqual(cantidad_etiquetas(filas), 3)
        salida = consolidar_etiquetas(filas)
        self.assertEqual(salida[1][:4], filas[1][:4])
        self.assertEqual(salida[1][4], "nombre:Ana;primer apellido:Pérez;segundo apellido:Gómez")
        self.assertEqual(salida[2][4], "nombre:Luis;primer apellido:Díaz;segundo apellido:")
        self.assertTrue(all(len(fila) == 5 for fila in salida))

    def test_sin_etiquetas_y_valores_cero(self):
        filas = [ENCABEZADOS, ["CC", "1", "NIT", "2"]]
        self.assertEqual(cantidad_etiquetas(filas), 0)
        self.assertEqual(consolidar_etiquetas(filas)[1][4], "")
        filas = [[*ENCABEZADOS, "importe"], ["CC", "1", "NIT", "2", 0]]
        self.assertEqual(consolidar_etiquetas(filas)[1][4], "importe:0")

    def test_etiquetas_sin_encabezado_se_reportan(self):
        filas = [[*ENCABEZADOS, "nombre", None], ["CC", "1", "NIT", "2", "Ana", "Pérez"]]
        errores = validar(filas)
        self.assertEqual(len(errores), 1)
        self.assertEqual((errores[0].fila, errores[0].columna), (1, 6))
        self.assertIn("Falta el encabezado", errores[0].motivo)


class ConversionTests(unittest.TestCase):
    def crear_base(self, ruta, filas):
        libro = Workbook(write_only=True)
        hoja = libro.create_sheet()
        for fila in filas:
            hoja.append(fila)
        libro.save(ruta)
        libro.close()

    def test_conversion_texto_vista_previa_y_nombres_sin_sobrescribir(self):
        from app import leer_vista_previa

        with tempfile.TemporaryDirectory() as carpeta:
            origen = Path(carpeta) / "base.xlsx"
            self.crear_base(origen, [
                [*ENCABEZADOS, "nombre", "primer apellido"],
                ["CC", "000123", "NIT", "12345678901234567890", "Ana María", "Pérez"],
            ])
            original = origen.read_bytes()
            salidas, errores = convertir(origen)
            self.assertEqual(errores, [])
            self.assertEqual(salidas[0].ruta.name, "base estructurado1.xlsx")
            self.assertEqual(salidas[0].registros, 1)
            filas = [list(f) for f in leer_excel(salidas[0].ruta)]
            self.assertEqual(filas[0], [*ENCABEZADOS, "parameters"])
            self.assertEqual(filas[1], ["CC", "000123", "NIT", "12345678901234567890", "nombre:Ana María;primer apellido:Pérez"])
            self.assertEqual([list(f) for f in leer_vista_previa(salidas[0].ruta)], filas)
            primera = salidas[0].ruta.read_bytes()
            repetidas, _ = convertir(origen)
            self.assertEqual(repetidas[0].ruta.name, "base estructurado1 (2).xlsx")
            self.assertEqual(salidas[0].ruta.read_bytes(), primera)
            self.assertEqual(origen.read_bytes(), original)

    def test_faltantes_documentos_estructura_y_celda_larga_bloquean_sin_csv(self):
        casos = [
            [[*ENCABEZADOS, "nombre"], ["CC", "1", "NIT", "2", None]],
            [[*ENCABEZADOS, None], ["CC", "1", "NIT", "2", "Ana"]],
            [ENCABEZADOS, ["CC", "A" * 21, "NIT", "2"]],
            [ENCABEZADOS, ["CC", "1", "", "2"]],
            [[*ENCABEZADOS[2:], *ENCABEZADOS[:2]], ["NIT", "2", "CC", "1"]],
            [[*ENCABEZADOS, "dato", "otro"], ["CC", "1", "NIT", "2", "A" * 17000, "B" * 17000]],
        ]
        for filas in casos:
            with self.subTest(filas=filas[:1]), tempfile.TemporaryDirectory() as carpeta:
                origen = Path(carpeta) / "base.xlsx"
                self.crear_base(origen, filas)
                salidas, errores = convertir(origen)
                self.assertEqual(salidas, [])
                self.assertTrue(errores)
                self.assertEqual(list(Path(carpeta).iterdir()), [origen])

    def test_limites_de_particion_y_resto(self):
        for total, esperados in [(1, [1]), (200000, [200000]), (200001, [200000, 1]),
                                 (250000, [200000, 50000]), (400000, [200000, 200000]),
                                 (480000, [200000, 200000, 80000])]:
            with self.subTest(total=total), tempfile.TemporaryDirectory() as carpeta:
                filas = [ENCABEZADOS] + [["CC", "1", "NIT", "2"]] * total
                tamanos = []
                def guardar(ruta, registros):
                    contenido = iter(registros)
                    self.assertEqual(tuple(next(contenido)), (*ENCABEZADOS, "parameters"))
                    tamanos.append(sum(1 for _ in contenido))
                    ruta.touch()
                with patch("validador.leer_excel", return_value=filas), patch("validador.guardar_xlsx", side_effect=guardar):
                    salidas, errores = convertir(Path(carpeta) / "base.xlsx")
                self.assertEqual(errores, [])
                self.assertEqual(tamanos, esperados)
                self.assertEqual([s.registros for s in salidas], esperados)
                self.assertEqual([s.ruta.name for s in salidas], [f"base estructurado{i}.xlsx" for i in range(1, len(esperados) + 1)])

    def test_particiones_reales_conservan_orden_encabezados_y_texto(self):
        with tempfile.TemporaryDirectory() as carpeta:
            origen = Path(carpeta) / "base.xlsx"
            self.crear_base(origen, [
                [*ENCABEZADOS, "nombre"],
                *[["CC", f"{i:05}", "NIT", "2", f"Persona {i}"] for i in range(5)],
            ])
            with patch("validador.REGISTROS_POR_ARCHIVO", 2):
                salidas, errores = convertir(origen)
            self.assertEqual(errores, [])
            self.assertEqual([s.registros for s in salidas], [2, 2, 1])
            registros = []
            for salida in salidas:
                filas = [list(f) for f in leer_excel(salida.ruta)]
                self.assertEqual(filas[0], [*ENCABEZADOS, "parameters"])
                registros.extend(filas[1:])
            self.assertEqual([r[1] for r in registros], [f"{i:05}" for i in range(5)])
            self.assertEqual([r[4] for r in registros], [f"nombre:Persona {i}" for i in range(5)])

    def test_error_en_ultima_parte_bloquea_todo_el_lote(self):
        with tempfile.TemporaryDirectory() as carpeta:
            origen = Path(carpeta) / "base.xlsx"
            self.crear_base(origen, [[*ENCABEZADOS, "nombre"],
                ["CC", "1", "NIT", "2", "Ana"], ["CC", "3", "NIT", "2", None]])
            with patch("validador.REGISTROS_POR_ARCHIVO", 1):
                salidas, errores = convertir(origen)
            self.assertFalse(salidas)
            self.assertTrue(errores)
            self.assertEqual(list(Path(carpeta).iterdir()), [origen])

    def test_valores_que_parecen_formulas_se_guardan_como_texto(self):
        from validador import guardar_xlsx
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "texto.xlsx"
            guardar_xlsx(ruta, [[*ENCABEZADOS, "parameters"], ["CC", "0001", "NIT", "2", "=texto:valor"]])
            libro = load_workbook(ruta)
            try:
                self.assertEqual(libro.active["E2"].data_type, "s")
                self.assertEqual(libro.active["B2"].value, "0001")
            finally:
                libro.close()


if __name__ == "__main__":
    unittest.main()
