import unittest

from excel_reader import EXCEL_FILE, load_sheet


class ExcelReaderTests(unittest.TestCase):
    def test_file_exists(self):
        self.assertTrue(EXCEL_FILE.exists(), "El archivo XLSX debe existir")

    def test_load_headers_and_rows(self):
        headers, rows = load_sheet()
        self.assertEqual(
            headers,
            ["ID", "Test Name", "Descripción", "Resultado Esperado", "Estado", "Evidencia"],
        )
        self.assertIsInstance(rows, list)
        self.assertEqual(len(rows), 0)


if __name__ == "__main__":
    unittest.main()
