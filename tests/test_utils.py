import unittest
import pandas as pd
import numpy as np
from consolidate_relatorio_base import (
    _canon_text,
    _normalize_column_name,
    _make_unique,
    _is_filled_cell,
    _looks_numeric,
    selecionar_engine_excel,
    SUPPORTED_EXCEL_EXTENSIONS,
)

class TestUtils(unittest.TestCase):
    
    def test_canon_text(self):
        self.assertEqual(_canon_text("  HELLO   WORLD  "), "hello world")
        self.assertEqual(_canon_text(None), "")
        self.assertEqual(_canon_text("Teste"), "teste")
        self.assertEqual(_canon_text("  "), "")

    def test_normalize_column_name(self):
        self.assertEqual(_normalize_column_name("  Nome  Completo  "), "Nome Completo")
        self.assertEqual(_normalize_column_name(None), "")
        self.assertEqual(_normalize_column_name(np.nan), "")
        self.assertEqual(_normalize_column_name("A" * 300), "A" * 255)

    def test_make_unique(self):
        self.assertEqual(_make_unique(["col", "col", "col"]), ["col", "col_2", "col_3"])
        self.assertEqual(_make_unique(["a", "b", "a"]), ["a", "b", "a_2"])
        self.assertEqual(_make_unique(["x", "y", "z"]), ["x", "y", "z"])

    def test_is_filled_cell(self):
        self.assertTrue(_is_filled_cell("texto"))
        self.assertTrue(_is_filled_cell(123))
        self.assertFalse(_is_filled_cell(None))
        self.assertFalse(_is_filled_cell(""))
        self.assertFalse(_is_filled_cell("  "))
        self.assertFalse(_is_filled_cell(np.nan))
        self.assertFalse(_is_filled_cell("NaN"))

    def test_looks_numeric(self):
        self.assertTrue(_looks_numeric("123.45"))
        self.assertTrue(_looks_numeric("1.234,56"))
        self.assertTrue(_looks_numeric("-123"))
        self.assertTrue(_looks_numeric("+123,45"))
        self.assertFalse(_looks_numeric("abc"))
        self.assertFalse(_looks_numeric("12.34.56"))
        self.assertFalse(_looks_numeric(None))

    def test_selecionar_engine_excel(self):
        casos = {
            "relatorio.xlsx": "openpyxl",
            "relatorio.XLSM": "openpyxl",
            "modelo.xltx": "openpyxl",
            "modelo.xltm": "openpyxl",
            "legado.xls": "xlrd",
            "binario.xlsb": "pyxlsb",
        }
        for arquivo, engine_esperada in casos.items():
            with self.subTest(arquivo=arquivo):
                self.assertEqual(selecionar_engine_excel(arquivo), engine_esperada)

    def test_formato_excel_nao_suportado(self):
        with self.assertRaisesRegex(ValueError, "Formato '.csv' não suportado"):
            selecionar_engine_excel("relatorio.csv")

    def test_lista_de_formatos_suportados(self):
        self.assertEqual(
            set(SUPPORTED_EXCEL_EXTENSIONS),
            {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb"},
        )

if __name__ == "__main__":
    unittest.main()
